"""
app.py — ProjectZen Middleware (Tier 2)
FastAPI server. All routes. Entry point.

Run with:
    uvicorn app:app --host 0.0.0.0 --port 8000 --reload

Routes:
    POST /generate          — generate xlsx/docx/pdf/pptx/xml
    POST /ask               — agentic Q&A grounded in document library
    POST /search            — direct keyword/filter search
    POST /upload-ref        — upload reference file
    GET  /download/{id}     — download a generated file
    GET  /documents         — list documents with filters
    GET  /ref/{ref_id}      — serve a reference file
    GET  /health            — health check
"""

import os
import re
import uuid
import json
import base64
import asyncio
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# Load .env from the middleware directory if it exists
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    try:
        from dotenv import load_dotenv
        load_dotenv(_env_path)
    except ImportError:
        # dotenv not installed — parse manually
        for _line in _env_path.read_text(encoding="utf-8").splitlines():
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _, _v = _line.partition("=")
                os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

import aiofiles
from fastapi import FastAPI, HTTPException, Request, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel
import anthropic

from db import (
    init_db, save_document, list_documents, search_documents, get_document,
    # cascade addon
    init_cascade_tables, seed_format_config_defaults,
    save_cascade_session, update_cascade_session_status,
    get_cascade_session, list_cascade_sessions,
    list_cascade_documents, finalise_cascade_document,
    get_cascade_doc_current_version,
    list_format_configs, update_format_config,
    get_cascade_doc_by_node_version,
    # grounding addon
    init_grounding_table, save_grounding_doc, get_grounding_doc,
    list_grounding_docs, delete_grounding_doc,
    # workbook multi-slot
    save_workbook_slot, get_workbook_slot, list_workbook_slots, delete_workbook_slot,
    MAX_WORKBOOK_SLOTS,
)
from extractor import extract_text, extract_from_bytes
from templates import generate, sanitize_xml_ref
from agent import ask_library
from knowledge_graph import (
    get_graph_data_for_frontend, get_all_node_ids, get_node,
    get_nodes_grouped_by_phase, compute_bfs_waves, get_downstream_nodes,
    DEFAULT_OUTPUT_FORMATS,
)

# ── Paths ─────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent
STORE_DIR  = BASE_DIR / "storage" / "document_store"
REFS_DIR   = BASE_DIR / "storage" / "document_store" / "refs"
STORE_DIR.mkdir(parents=True, exist_ok=True)
REFS_DIR.mkdir(parents=True, exist_ok=True)

MODEL    = "claude-opus-4-8"
MAX_TOKS = 32000

# ── App ───────────────────────────────────────────────────────────────────
app = FastAPI(title="ProjectZen Middleware", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup():
    await init_db()
    await init_cascade_tables()
    await seed_format_config_defaults()
    await init_grounding_table()
    print(f"🚀 ProjectZen Middleware started")
    print(f"   Store : {STORE_DIR}")
    print(f"   Refs  : {REFS_DIR}")
    print(f"   Model : {MODEL}")


# ═══════════════════════════════════════════════════════════════════════════
# Pydantic models
# ═══════════════════════════════════════════════════════════════════════════

class FilePayload(BaseModel):
    type:    str                    # "base64" | "text" | "refId"
    ext:     Optional[str] = None
    content: Optional[str] = None   # base64 string or raw text
    name:    Optional[str] = None
    mime:    Optional[str] = None
    refId:   Optional[str] = None


class GenerateRequest(BaseModel):
    userPrompt:     str
    outputFormat:   str = "xlsx"
    fileName:       Optional[str] = None
    fileData:       Optional[FilePayload] = None
    refData:        Optional[FilePayload] = None
    mirrorAspects:  List[str] = []
    assistantMemory: Optional[str] = None
    isRefinement:   bool = False
    username:       Optional[str] = "unknown"


class AskRequest(BaseModel):
    question: str
    username: Optional[str] = None


class SearchRequest(BaseModel):
    query:        str
    format:       Optional[str] = None
    username:     Optional[str] = None
    date_from:    Optional[str] = None
    date_to:      Optional[str] = None
    limit:        int = 20


class UploadRefRequest(BaseModel):
    content:  str
    name:     str
    ext:      str
    mime:     Optional[str] = None
    refType:  str = "base64"


# ═══════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════

def safe_cell(v: Any, maxlen: int = 100) -> str:
    if v is None:
        return ""
    s = str(v)
    s = "".join(c if 32 <= ord(c) < 127 else " " for c in s)
    return s[:maxlen].strip()


def sanitize_for_json(text: str) -> str:
    if not text:
        return ""
    s = "".join(c if 32 <= ord(c) < 127 or c == "\n" else " " for c in text)
    s = s.replace('"', "'").replace("\\", " ")
    s = re.sub(r"[\r\n]+", " | ", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()[:40000]


def sanitize_plan(obj: Any) -> Any:
    if isinstance(obj, list):
        return [sanitize_plan(i) for i in obj]
    if isinstance(obj, dict):
        return {k: sanitize_plan(v) for k, v in obj.items()}
    if isinstance(obj, str):
        return safe_cell(obj)
    if isinstance(obj, float):
        return obj if obj == obj else 0   # NaN check
    return obj


def clean_json(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*```\s*$", "", s)
    start = s.find("{")
    end   = s.rfind("}")
    if start != -1 and end > start:
        s = s[start:end + 1]
    s = re.sub(r",(\s*[}\]])", r"\1", s)
    return s


def make_file_name(template_name: str, output_format: str) -> str:
    safe = re.sub(r"\s*→\s*", "_to_", template_name or "Document")
    safe = re.sub(r"[\s]+", "_", safe)
    safe = re.sub(r"[()/:*?\"<>|]", "", safe)
    safe = re.sub(r"_+", "_", safe).strip("_")[:60]
    now  = datetime.utcnow()
    ts   = now.strftime("%Y-%m-%d_%H-%M-%S")
    ext  = output_format.lower()
    return f"{safe}_{ts}.{ext}"


async def resolve_file(fd: Optional[FilePayload]) -> str:
    """Resolve a FilePayload to extracted text."""
    if not fd:
        return ""
    if fd.type == "text":
        return (fd.content or "")[:80000]
    if fd.type == "base64" and fd.content and fd.ext:
        raw = base64.b64decode(fd.content)
        return extract_from_bytes(raw, fd.ext)
    if fd.type == "refId" and fd.refId:
        files = list(REFS_DIR.glob(f"{fd.refId}*"))
        if not files:
            print(f"   ⚠ refId not found: {fd.refId}")
            return ""
        return extract_text(str(files[0]))
    return ""


# ── FIX C: XML prompt labelling ───────────────────────────────────────────

def build_json_prompt(
    user_prompt: str,
    output_format: str,
    safe_input: str,
    safe_ref: str,
    mirror_aspects: List[str],
    attempt: int,
    is_refinement: bool = False,
    raw_ref_text: str = "",
) -> str:
    parts   = user_prompt.split("\n\nAdditional user instructions:\n")
    sys_p   = f"SYSTEM INSTRUCTIONS:\n{parts[0][:3000]}\n" if parts[0] else ""
    usr_p   = f"USER PERSONALIZATION:\n{parts[1][:1500]}\n" if len(parts) > 1 else ""
    req_blk = (sys_p + usr_p) or f"REQUEST: {user_prompt[:4000]}\n"

    input_blk = f"\nINPUT CONTENT (extract all details):\n{safe_input}\n" if safe_input else ""

    # FIX C: XML reference gets explicit schema label
    if output_format == "xml" and raw_ref_text:
        schema_snippet = sanitize_xml_ref(raw_ref_text)
        ref_blk = (
            f"\nEXACT XML SCHEMA TO FOLLOW:\n{schema_snippet}\n\n"
            "CRITICAL XML INSTRUCTIONS:\n"
            "- Use the SAME root element tag name as the reference\n"
            "- Use the SAME entity/field tag names as the reference\n"
            "- Populate Records with data from INPUT CONTENT\n"
        )
    elif safe_ref:
        ref_blk = f"\nREFERENCE STYLE:\n{safe_ref[:8000]}\n"
    else:
        ref_blk = ""

    mirror  = f"Mirror these aspects: {', '.join(mirror_aspects)}." if mirror_aspects else ""
    warn    = "\n!! PREVIOUS ATTEMPT PRODUCED INVALID JSON. Be extra careful.\n" if attempt > 1 else ""
    refine  = "\nREFINEMENT MODE: Extend/modify the previous document. Output COMPLETE JSON.\n" if is_refinement else ""

    schemas = {
        "xml":  '{"title":"string","subtitle":"string","sheets":[{"name":"string","headers":["string"],"rows":[["string"]],"has_totals":false}],"sections":[{"heading":"string","paragraphs":["string"],"bullets":["string"],"table":{"headers":["string"],"rows":[["string"]]}}]}',
        "xlsx": '{"title":"string","sheets":[{"name":"string","headers":["string"],"rows":[["string"]],"has_totals":false}],"chart_sheet":0,"chart_data_col":1}',
        "docx": '{"title":"string","subtitle":"string","sections":[{"heading":"string","level":1,"paragraphs":["string"],"bullets":["string"],"table":{"headers":["string"],"rows":[["string"]]}}]}',
        "pdf":  '{"title":"string","subtitle":"string","sections":[{"heading":"string","paragraphs":["string"],"table":{"headers":["string"],"rows":[["string"]]}}]}',
        "pptx": '{"title":"string","subtitle":"string","slides":[{"type":"title","title":"string","subtitle":"string"},{"type":"bullets","title":"string","bullets":["string"]},{"type":"table","title":"string","headers":["string"],"rows":[["string"]]},{"type":"chart","title":"string","categories":["string"],"values":[1,2,3],"series_name":"string"}]}',
    }

    detail = {
        "xml":  "Generate AT LEAST 3-5 entities. Each entity: 8-15 records. Use reference field names exactly.",
        "xlsx": "Generate AT LEAST 5-7 sheets. Each sheet: 10-25 data rows. Extract real tasks, dates, owners from input.",
        "docx": "Generate AT LEAST 6-10 sections. Each: 2-4 paragraphs. Tables: 8-15 rows.",
        "pdf":  "Generate AT LEAST 6-10 sections. Each: 2-4 paragraphs.",
        "pptx": "Generate AT LEAST 10-15 slides. Mix title, bullets, table, chart types.",
    }

    return (
        "Return ONLY a valid JSON object. No markdown. No explanation. No backticks.\n"
        + warn + refine + "\n"
        + req_blk + "\n"
        + mirror + "\n"
        + input_blk + "\n"
        + ref_blk + "\n"
        + f"\nSchema for {output_format.upper()} (follow exactly):\n"
        + schemas.get(output_format, schemas["docx"]) + "\n\n"
        + detail.get(output_format, detail["docx"]) + "\n\n"
        + "HARD RULES:\n"
        + "1. Every string value: max 80 chars, ASCII only.\n"
        + "2. No trailing commas.\n"
        + "3. Numbers must be JSON numbers in values arrays.\n"
        + "4. Return ONLY the JSON object."
    )


async def get_content_plan(
    user_prompt: str,
    output_format: str,
    input_text: str,
    ref_text: str,
    mirror_aspects: List[str],
    assistant_memory: Optional[str],
    is_refinement: bool,
    raw_ref_text: str,
) -> Dict:
    client     = anthropic.AsyncAnthropic()
    safe_input = sanitize_for_json(input_text)
    safe_ref   = "" if output_format == "xml" else sanitize_for_json(ref_text)

    for attempt in range(1, 4):
        print(f"   → Plan attempt {attempt}/3 ({MODEL})...")
        prompt = build_json_prompt(
            user_prompt, output_format, safe_input, safe_ref,
            mirror_aspects, attempt, is_refinement, raw_ref_text,
        )

        messages: List[Dict] = []
        if assistant_memory:
            messages += [
                {"role": "user",      "content": "You previously generated a document. Summary:"},
                {"role": "assistant", "content": assistant_memory},
                {"role": "user",      "content": prompt},
            ]
        else:
            messages = [{"role": "user", "content": prompt}]

        raw = ""
        try:
            async with client.messages.stream(
                model=MODEL, max_tokens=MAX_TOKS, messages=messages
            ) as stream:
                async for text in stream.text_stream:
                    raw += text
        except Exception as api_err:
            print(f"   → API error: {str(api_err)[:200]}")
            if attempt == 3:
                raise
            continue

        try:
            plan = json.loads(clean_json(raw))
            return sanitize_plan(plan)
        except Exception as parse_err:
            print(f"   → Parse failed (attempt {attempt}): {str(parse_err)[:100]}")

    # Fallback plan
    return {
        "title": user_prompt[:50],
        "sheets": [{"name": "Data", "headers": ["Item", "Details"], "rows": [["1", "See source"]], "has_totals": False}],
        "sections": [{"heading": "Summary", "level": 1, "paragraphs": ["Document generated."], "bullets": [], "table": None}],
    }


# ═══════════════════════════════════════════════════════════════════════════
# Routes
# ═══════════════════════════════════════════════════════════════════════════

@app.post("/generate")
async def route_generate(req: GenerateRequest, request: Request):
    if not req.userPrompt:
        raise HTTPException(400, "userPrompt is required")

    print(f"\n📄 /generate | format={req.outputFormat} | input={req.fileData.name if req.fileData else 'none'}")

    # ── Detect CSF XML: Excel input + XML output + CSF keyword in prompt/template ──
    is_csf = (
        req.outputFormat == "xml"
        and req.fileData is not None
        and req.fileData.ext and req.fileData.ext.lower() in ("xlsx", "xls")
        and any(kw in (req.userPrompt + (req.fileName or "")).lower()
                for kw in ("csf", "country specific", "country-specific",
                           "successfactor", "data model csf", "hris"))
    )

    if is_csf:
        return await _handle_csf_xml(req, request)

    # ── Standard generation flow ───────────────────────────────────────────
    input_text   = await resolve_file(req.fileData)
    raw_ref_text = ""
    ref_text     = ""

    if req.refData:
        if req.outputFormat == "xml":
            raw_ref_text = await resolve_file(req.refData)
        else:
            ref_text = await resolve_file(req.refData)

    plan = await get_content_plan(
        req.userPrompt, req.outputFormat, input_text, ref_text,
        req.mirrorAspects, req.assistantMemory, req.isRefinement, raw_ref_text,
    )

    ref_schema = ""
    if req.outputFormat == "xml" and raw_ref_text:
        plan["refSchema"] = raw_ref_text[:20000]
        ref_schema        = raw_ref_text

    template_match = re.search(r"^Template:\s*(.+)$", req.userPrompt, re.MULTILINE)
    template_name  = template_match.group(1).strip() if template_match else (req.fileName or plan.get("title", "Document"))
    out_name       = make_file_name(template_name, req.outputFormat)
    out_path       = str(STORE_DIR / out_name)

    try:
        generate(req.outputFormat, plan, out_path, ref_schema)
    except Exception as gen_err:
        raise HTTPException(500, f"Generation failed: {str(gen_err)[:400]}")

    size_kb = round(os.path.getsize(out_path) / 1024, 1)

    content_summary = ""
    try:
        content_summary = extract_text(out_path, req.outputFormat)[:8000]
    except Exception:
        content_summary = f"Title: {plan.get('title', template_name)}"

    doc_id = str(uuid.uuid4())
    await save_document(
        id             = doc_id,
        file_name      = out_name,
        template       = template_name,
        output_format  = req.outputFormat,
        file_path      = out_path,
        size_kb        = size_kb,
        username       = req.username or "unknown",
        content_summary= content_summary,
    )

    base_url = str(request.base_url).rstrip("/")
    return {
        "success":      True,
        "fileId":       doc_id,
        "fileName":     out_name,
        "outputFormat": req.outputFormat,
        "sizeKB":       size_kb,
        "downloadUrl":  f"{base_url}/download/{doc_id}?name={out_name}",
        "preview":      f"✅ File ready!\nFile: {out_name}\nSize: {size_kb} KB | Format: {req.outputFormat.upper()}",
    }


async def _handle_csf_xml(req: GenerateRequest, request: Request):
    """
    Dedicated handler for Excel → CSF XML conversion.
    Bypasses Claude entirely — reads CSF sheets directly from the workbook
    and produces the exact SAP SF country-specific-fields XML structure.
    """
    import tempfile, base64

    print(f"   🌍 CSF XML route detected — bypassing Claude, reading CSF sheets directly")

    # Write the uploaded Excel to a temp file
    if req.fileData.type == "base64" and req.fileData.content:
        raw = base64.b64decode(req.fileData.content)
        suffix = f".{req.fileData.ext or 'xlsx'}"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(raw)
            excel_tmp = tmp.name
    elif req.fileData.type == "refId" and req.fileData.refId:
        files = list(REFS_DIR.glob(f"{req.fileData.refId}*"))
        if not files:
            raise HTTPException(400, "Reference file not found")
        excel_tmp = str(files[0])
    else:
        raise HTTPException(400, "Excel file content required for CSF XML generation")

    template_match = re.search(r"^Template:\s*(.+)$", req.userPrompt, re.MULTILINE)
    template_name  = template_match.group(1).strip() if template_match else (req.fileName or "CSF_DataModel")
    out_name       = make_file_name(template_name, "xml")
    out_path       = str(STORE_DIR / out_name)

    try:
        from templates import generate_csf_xml
        generate_csf_xml(excel_tmp, out_path)
    except Exception as gen_err:
        raise HTTPException(500, f"CSF XML generation failed: {str(gen_err)[:400]}")
    finally:
        # Clean up temp file if we created it
        if req.fileData.type == "base64":
            try:
                os.unlink(excel_tmp)
            except Exception:
                pass

    size_kb = round(os.path.getsize(out_path) / 1024, 1)

    # Read a snippet for content_summary
    try:
        with open(out_path, "r", encoding="utf-8") as f:
            content_summary = f.read(8000)
    except Exception:
        content_summary = f"CSF XML: {template_name}"

    doc_id = str(uuid.uuid4())
    await save_document(
        id             = doc_id,
        file_name      = out_name,
        template       = template_name,
        output_format  = "xml",
        file_path      = out_path,
        size_kb        = size_kb,
        username       = req.username or "unknown",
        content_summary= content_summary,
    )

    base_url = str(request.base_url).rstrip("/")
    return {
        "success":      True,
        "fileId":       doc_id,
        "fileName":     out_name,
        "outputFormat": "xml",
        "sizeKB":       size_kb,
        "downloadUrl":  f"{base_url}/download/{doc_id}?name={out_name}",
        "preview":      (
            f"✅ CSF XML ready!\nFile: {out_name}\n"
            f"Size: {size_kb} KB | SAP SuccessFactors country-specific-fields format"
        ),
    }


@app.post("/ask")
async def route_ask(req: AskRequest, request: Request):
    """Agentic Q&A grounded in document library."""
    if not req.question.strip():
        raise HTTPException(400, "question is required")

    base_url = str(request.base_url).rstrip("/")
    result   = await ask_library(req.question, req.username, base_url)

    return {
        "success":        True,
        "answer":         result["answer"],
        "documents":      result["documents"],
        "toolCallsMade":  result["tool_calls_made"],
    }


@app.post("/search")
async def route_search(req: SearchRequest):
    """Direct keyword/filter search — no agent loop."""
    results = await search_documents(
        query         = req.query,
        output_format = req.format,
        username      = req.username,
        date_from     = req.date_from,
        date_to       = req.date_to,
        limit         = req.limit,
    )
    return {"success": True, "count": len(results), "documents": results}


@app.post("/upload-ref")
async def route_upload_ref(req: UploadRefRequest):
    """Store an uploaded reference file and return its refId."""
    ref_id    = str(uuid.uuid4())
    safe_ext  = re.sub(r"[^a-zA-Z0-9]", "", req.ext)
    file_name = f"{ref_id}.{safe_ext}"
    file_path = REFS_DIR / file_name

    if req.refType == "text":
        async with aiofiles.open(file_path, "w", encoding="utf-8") as f:
            await f.write(req.content)
    else:
        raw = base64.b64decode(req.content)
        async with aiofiles.open(file_path, "wb") as f:
            await f.write(raw)

    size_kb = round(os.path.getsize(file_path) / 1024, 1)
    print(f"   📎 Ref saved: {req.name} ({size_kb} KB) → {file_name}")

    return {
        "success": True,
        "refId":   ref_id,
        "fileName": req.name,
        "ext":      safe_ext,
        "sizeKB":   size_kb,
    }


@app.get("/download/{file_id}")
async def route_download(file_id: str, name: Optional[str] = Query(None)):
    """Download a generated file by its document ID."""
    doc = await get_document(file_id)
    if not doc:
        raise HTTPException(404, "Document not found")

    file_path = doc["file_path"]
    if not os.path.exists(file_path):
        raise HTTPException(404, "File not found on disk")

    return FileResponse(
        path            = file_path,
        filename        = name or doc["file_name"],
        media_type      = "application/octet-stream",
    )


@app.get("/documents")
async def route_documents(
    username: Optional[str] = None,
    format:   Optional[str] = None,
    template: Optional[str] = None,
    limit:    int = 100,
):
    docs = await list_documents(username=username, output_format=format, template=template, limit=limit)
    return {"success": True, "count": len(docs), "documents": docs}


@app.get("/ref/{ref_id}")
async def route_ref(ref_id: str):
    """Serve a reference file."""
    files = list(REFS_DIR.glob(f"{ref_id}*"))
    if not files:
        raise HTTPException(404, "Reference file not found")
    return FileResponse(path=str(files[0]), media_type="application/octet-stream")


@app.get("/health")
async def route_health():
    from db import DB_PATH
    import aiosqlite
    doc_count = 0
    try:
        async with aiosqlite.connect(DB_PATH) as db:
            async with db.execute("SELECT COUNT(*) FROM documents") as cur:
                row = await cur.fetchone()
                doc_count = row[0] if row else 0
    except Exception:
        pass

    return {
        "status":           "ok",
        "model":            MODEL,
        "version":          "1.0.0",
        "documents_stored": doc_count,
        "store_dir":        str(STORE_DIR),
        "timestamp":        datetime.utcnow().isoformat(),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# CASCADE ADDON ROUTES  (user login only — enforced by frontend)
# ═══════════════════════════════════════════════════════════════════════════════

# ── Pydantic models for cascade ───────────────────────────────────────────────

class CascadeStartRequest(BaseModel):
    fileData:      FilePayload
    selectedNodes: List[str] = []      # empty = all 32
    rootNodeId:    Optional[str] = None  # root/input doc node ID (excluded from generation)
    username:      Optional[str] = "unknown"


class DeltaAnalyseRequest(BaseModel):
    sessionId:     str                 # Source session containing the old version
    nodeId:        str                 # Which document changed
    fromVersion:   int = 1             # Version to compare against
    inputType:     str = "full_doc"    # "full_doc" | "delta_section" | "prompt"
    fileData:      Optional[FilePayload] = None   # For full_doc / delta_section
    promptText:    Optional[str] = None           # For prompt mode
    username:      Optional[str] = "unknown"


class DeltaApplyRequest(BaseModel):
    sourceSessionId:      str
    deltaSessionId:       str
    deltaNodeId:          str
    selectedNodes:        List[str]              # Nodes user confirmed to update
    finalisedDecisions:   Dict[str, bool] = {}   # node_id -> True = include
    username:             Optional[str] = "unknown"


class FinaliseDocRequest(BaseModel):
    sessionId: str
    nodeId:    str
    version:   int
    username:  Optional[str] = "unknown"


class FormatConfigUpdateRequest(BaseModel):
    nodeId:        str
    outputFormat:  str
    isDefaultSel:  int = 1
    username:      Optional[str] = "admin"


class GroundingUploadRequest(BaseModel):
    nodeId:   str
    content:  str            # base64-encoded file bytes
    name:     str            # original file name shown to user
    ext:      str            # file extension (xlsx / docx / pdf / xml)
    mime:     Optional[str] = None
    username: Optional[str] = "admin"


# ── GET /cascade/graph-data ───────────────────────────────────────────────────

@app.get("/cascade/graph-data")
async def route_cascade_graph_data():
    """Return the full knowledge graph data for the frontend Cytoscape.js visualization."""
    return {"success": True, "data": get_graph_data_for_frontend()}


# ── POST /cascade/start ───────────────────────────────────────────────────────

@app.post("/cascade/start")
async def route_cascade_start(req: CascadeStartRequest, request: Request):
    """
    Start a new cascade generation session.
    Extracts text from the uploaded document, then fires the cascade in the background.
    Returns session_id immediately — client connects to /cascade/stream/{session_id} for SSE.
    """
    if not req.fileData:
        raise HTTPException(400, "fileData is required")

    # Extract input document text
    input_text = await resolve_file(req.fileData)
    if not input_text.strip():
        raise HTTPException(400, "Could not extract text from the uploaded document")

    # Determine selected nodes
    selected = req.selectedNodes if req.selectedNodes else get_all_node_ids()
    # Validate node ids
    valid_ids  = set(get_all_node_ids())
    selected   = [n for n in selected if n in valid_ids]
    if not selected:
        raise HTTPException(400, "No valid document nodes selected")

    session_id = str(uuid.uuid4())
    file_name  = req.fileData.name or "input_document"

    # Set up SSE queue before launching background task
    from context_store import create_sse_queue
    create_sse_queue(session_id)

    # Launch cascade in background (non-blocking)
    from cascade_agent import run_new_cascade
    asyncio.create_task(run_new_cascade(
        session_id=session_id,
        input_text=input_text,
        input_file_name=file_name,
        selected_nodes=selected,
        root_node_id=req.rootNodeId or None,
        username=req.username or "unknown",
    ))

    return {
        "success":    True,
        "sessionId":  session_id,
        "totalDocs":  len(selected),
        "streamUrl":  f"/cascade/stream/{session_id}",
        "message":    f"Cascade started for {len(selected)} documents.",
    }


# ── GET /cascade/stream/{session_id} — SSE endpoint ──────────────────────────

@app.get("/cascade/stream/{session_id}")
async def route_cascade_stream(session_id: str):
    """
    Server-Sent Events stream for a cascade session.
    Drains the asyncio.Queue set up by /cascade/start.
    Each event: data: {json}\n\n
    """
    from context_store import get_sse_queue, cleanup_sse_queue

    async def event_generator():
        q = get_sse_queue(session_id)
        if not q:
            yield f"data: {json.dumps({'event': 'error', 'data': {'message': 'Session not found'}})}\n\n"
            return

        try:
            while True:
                try:
                    item = await asyncio.wait_for(q.get(), timeout=60.0)
                except asyncio.TimeoutError:
                    # Keep-alive ping
                    yield f": ping\n\n"
                    continue

                if item is None:
                    yield f"data: {json.dumps({'event': 'stream_end', 'data': {}})}\n\n"
                    break

                event_type = item.get("event", "message")
                data       = item.get("data", {})
                yield f"data: {json.dumps({'event': event_type, 'data': data})}\n\n"

        except asyncio.CancelledError:
            pass
        finally:
            cleanup_sse_queue(session_id)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control":               "no-cache",
            "X-Accel-Buffering":           "no",
            "Access-Control-Allow-Origin": "*",
        },
    )


# ── POST /cascade/delta/analyse ───────────────────────────────────────────────

@app.post("/cascade/delta/analyse")
async def route_delta_analyse(req: DeltaAnalyseRequest):
    """
    Run Agent 4 (Delta Analysis) to identify what changed and downstream impact.
    Returns delta_items and affected_nodes for the user to review and confirm.
    This is synchronous (waits for analysis to complete, typically 10-30s).
    """
    # Resolve new document text
    new_text = ""
    if req.inputType == "prompt":
        new_text = req.promptText or ""
    elif req.fileData:
        new_text = await resolve_file(req.fileData)

    if not new_text.strip():
        raise HTTPException(400, "No change content provided (prompt, file, or delta section required)")

    from delta_engine import analyse_delta
    result = await analyse_delta(
        session_id=req.sessionId,
        node_id=req.nodeId,
        from_version=req.fromVersion,
        new_content=new_text,
        input_type=req.inputType,
    )

    return {
        "success":       True,
        "deltaItems":    result.delta_items,
        "affectedNodes": result.affected_nodes,
        "deltaSummary":  result.delta_summary,
    }


# ── POST /cascade/delta/apply ─────────────────────────────────────────────────

@app.post("/cascade/delta/apply")
async def route_delta_apply(req: DeltaApplyRequest, request: Request):
    """
    After user confirms scope, start the delta cascade.
    Returns a new session_id for the delta run — client streams from /cascade/stream/{id}.
    """
    if not req.selectedNodes:
        raise HTTPException(400, "No nodes selected for delta update")

    # Load source session's project context and generated docs
    source_session = await get_cascade_session(req.sourceSessionId)
    if not source_session:
        raise HTTPException(404, f"Source session '{req.sourceSessionId}' not found")

    # Load project context from the LangGraph checkpointer state
    # Fallback: rebuild minimal context from session metadata
    from cascade_agent import _checkpointer
    config = {"configurable": {"thread_id": req.sourceSessionId}}
    try:
        snap = await _checkpointer.aget(config)
        state_vals = snap.values if snap else {}
    except Exception:
        state_vals = {}

    project_context = state_vals.get("project_context", {"project_name": "Project", "client_name": "Client", "delivery_partner": "Accenture"})
    generated_docs  = state_vals.get("generated_docs", {})

    # Compute delta waves for selected nodes
    delta_waves = compute_bfs_waves(req.selectedNodes)

    # Create new session_id for the delta run
    delta_session_id = str(uuid.uuid4())

    # Retrieve delta_items from the analyse step (caller must pass them)
    # For now we generate minimal delta_items placeholder
    delta_items   = []
    delta_summary = f"Delta update to {req.deltaNodeId}"

    # Set up SSE queue
    from context_store import create_sse_queue
    create_sse_queue(delta_session_id)

    from cascade_agent import run_delta_cascade
    asyncio.create_task(run_delta_cascade(
        session_id=delta_session_id,
        source_session_id=req.sourceSessionId,
        delta_node_id=req.deltaNodeId,
        delta_waves=delta_waves,
        delta_items=delta_items,
        delta_summary=delta_summary,
        user_delta_nodes=req.selectedNodes,
        user_finalised_choices=req.finalisedDecisions,
        project_context=project_context,
        generated_docs=generated_docs,
        username=req.username or "unknown",
    ))

    return {
        "success":        True,
        "deltaSessionId": delta_session_id,
        "streamUrl":      f"/cascade/stream/{delta_session_id}",
        "totalNodes":     len(req.selectedNodes),
        "deltaWaves":     len(delta_waves),
    }


# ── GET /cascade/session/{session_id} ─────────────────────────────────────────

@app.get("/cascade/session/{session_id}")
async def route_cascade_session(session_id: str, request: Request):
    """Return session metadata + all generated documents."""
    session = await get_cascade_session(session_id)
    if not session:
        raise HTTPException(404, "Session not found")

    docs     = await list_cascade_documents(session_id)
    base_url = str(request.base_url).rstrip("/")

    # Enrich docs with download URLs
    for d in docs:
        d["download_url"] = f"{base_url}/download/{d['file_id']}?name={d['file_name']}"

    return {"success": True, "session": session, "documents": docs}


# ── GET /cascade/sessions ─────────────────────────────────────────────────────

@app.get("/cascade/sessions")
async def route_cascade_sessions(username: Optional[str] = None, limit: int = 50):
    sessions = await list_cascade_sessions(username=username, limit=limit)
    return {"success": True, "count": len(sessions), "sessions": sessions}


# ── POST /cascade/finalise ────────────────────────────────────────────────────

@app.post("/cascade/finalise")
async def route_finalise_doc(req: FinaliseDocRequest):
    """Mark a specific document version as finalised (client signed-off)."""
    ok = await finalise_cascade_document(
        session_id=req.sessionId,
        node_id=req.nodeId,
        version=req.version,
        finalised_by=req.username or "unknown",
    )
    if not ok:
        raise HTTPException(500, "Failed to finalise document")
    return {"success": True, "message": f"{req.nodeId} v{req.version} marked as finalised"}


# ── GET /cascade/format-config ────────────────────────────────────────────────

@app.get("/cascade/format-config")
async def route_get_format_config():
    """Return admin-configurable output format settings for all document types."""
    configs = await list_format_configs()
    return {"success": True, "configs": configs}


# ── PUT /cascade/format-config ────────────────────────────────────────────────

@app.put("/cascade/format-config")
async def route_update_format_config(req: FormatConfigUpdateRequest):
    """Admin: update output format and default selection for a document type."""
    valid_formats = {"xlsx", "docx", "pdf", "pptx", "xml"}
    if req.outputFormat not in valid_formats:
        raise HTTPException(400, f"outputFormat must be one of {valid_formats}")
    valid_nodes = set(get_all_node_ids())
    if req.nodeId not in valid_nodes:
        raise HTTPException(400, f"nodeId '{req.nodeId}' not recognised")

    ok = await update_format_config(
        node_id=req.nodeId,
        output_format=req.outputFormat,
        is_default_sel=req.isDefaultSel,
        updated_by=req.username or "admin",
    )
    if not ok:
        raise HTTPException(500, "Failed to update format config")
    return {"success": True, "message": f"{req.nodeId} → {req.outputFormat} saved"}


# ═══════════════════════════════════════════════════════════════════════════════
# ADMIN GROUNDING ROUTES  (admin login only — enforced by frontend)
# ═══════════════════════════════════════════════════════════════════════════════

@app.get("/admin/grounding")
async def route_admin_grounding_list():
    """
    Return all 32 knowledge graph nodes with their grounding document status.
    Used by the admin Document Grounding panel to colour-code the graph.
    """
    from knowledge_graph import get_all_node_ids, get_node
    docs = await list_grounding_docs()
    grounded_map = {d["node_id"]: d for d in docs}

    nodes = []
    for node_id in get_all_node_ids():
        node      = get_node(node_id)
        grounding = grounded_map.get(node_id)
        nodes.append({
            "node_id":    node_id,
            "label":      node["label"] if node else node_id,
            "phase":      node["phase"] if node else "",
            "tier":       node["tier"]  if node else "",
            "is_grounded": grounding is not None,
            "grounding":   grounding,
        })
    grounded_count = sum(1 for n in nodes if n["is_grounded"])
    return {
        "success":       True,
        "nodes":         nodes,
        "total":         len(nodes),
        "grounded_count": grounded_count,
    }


@app.post("/admin/grounding/upload")
async def route_admin_grounding_upload(req: GroundingUploadRequest):
    """
    Upload a grounding reference document for a specific knowledge graph node.
    Replaces any existing grounding doc for that node.
    File is saved to REFS_DIR; record is stored in grounding_docs table.
    """
    from knowledge_graph import get_all_node_ids
    valid_nodes = set(get_all_node_ids())
    if req.nodeId not in valid_nodes:
        raise HTTPException(400, f"nodeId '{req.nodeId}' not recognised")

    safe_ext  = re.sub(r"[^a-zA-Z0-9]", "", req.ext)
    ref_id    = str(uuid.uuid4())
    file_path = REFS_DIR / f"{ref_id}.{safe_ext}"

    try:
        raw = base64.b64decode(req.content)
    except Exception:
        raise HTTPException(400, "Invalid base64 content")

    async with aiofiles.open(file_path, "wb") as f:
        await f.write(raw)

    size_kb = round(os.path.getsize(file_path) / 1024, 1)

    # Remove old grounding file from disk before replacing DB record
    old = await get_grounding_doc(req.nodeId)
    if old:
        for old_file in REFS_DIR.glob(f"{old['ref_id']}*"):
            try:
                os.unlink(old_file)
            except Exception:
                pass

    ok = await save_grounding_doc(
        node_id     = req.nodeId,
        ref_id      = ref_id,
        file_name   = req.name,
        file_ext    = safe_ext,
        size_kb     = size_kb,
        uploaded_by = req.username or "admin",
    )
    if not ok:
        raise HTTPException(500, "Failed to save grounding document record")

    print(f"   🔗 Grounding doc saved: {req.nodeId} ← {req.name} ({size_kb} KB)")
    return {
        "success":  True,
        "nodeId":   req.nodeId,
        "refId":    ref_id,
        "fileName": req.name,
        "sizeKB":   size_kb,
    }


@app.delete("/admin/grounding/{node_id}")
async def route_admin_grounding_delete(node_id: str):
    """Remove the grounding document for a knowledge graph node."""
    grounding = await get_grounding_doc(node_id)
    if not grounding:
        raise HTTPException(404, f"No grounding document found for node '{node_id}'")

    for old_file in REFS_DIR.glob(f"{grounding['ref_id']}*"):
        try:
            os.unlink(old_file)
        except Exception:
            pass

    ok = await delete_grounding_doc(node_id)
    if not ok:
        raise HTTPException(500, "Failed to remove grounding document record")

    print(f"   🗑 Grounding doc removed: {node_id}")
    return {"success": True, "message": f"Grounding document removed for '{node_id}'"}


# ═══════════════════════════════════════════════════════════════════════════════
# WORKBOOK MULTI-SLOT ENDPOINTS
# Configuration Workbook node supports up to 4 grounding reference workbooks
# ═══════════════════════════════════════════════════════════════════════════════

class WorkbookSlotUploadRequest(BaseModel):
    slot:     int             # 1–4
    content:  str             # base64-encoded xlsx bytes
    name:     str             # original file name
    ext:      str = "xlsx"
    username: Optional[str] = "admin"


@app.get("/admin/workbook-slots")
async def route_list_workbook_slots():
    """Return status of all 4 configuration workbook grounding slots."""
    occupied = {r["slot"]: r for r in await list_workbook_slots()}
    slots = []
    for i in range(1, MAX_WORKBOOK_SLOTS + 1):
        if i in occupied:
            s = occupied[i]
            slots.append({
                "slot": i,
                "occupied": True,
                "file_name": s["file_name"],
                "file_ext": s["file_ext"],
                "ref_id": s["ref_id"],
                "size_kb": s["size_kb"],
                "uploaded_by": s["uploaded_by"],
                "uploaded_at": s["uploaded_at"],
            })
        else:
            slots.append({"slot": i, "occupied": False})
    return {"success": True, "slots": slots}


@app.post("/admin/workbook-slots/upload")
async def route_upload_workbook_slot(req: WorkbookSlotUploadRequest):
    """Upload a reference workbook for one of the 4 config-workbook grounding slots."""
    if req.slot < 1 or req.slot > MAX_WORKBOOK_SLOTS:
        raise HTTPException(400, f"slot must be 1–{MAX_WORKBOOK_SLOTS}")

    safe_ext  = re.sub(r"[^a-zA-Z0-9]", "", req.ext) or "xlsx"
    ref_id    = str(uuid.uuid4())
    file_path = REFS_DIR / f"{ref_id}.{safe_ext}"

    try:
        raw = base64.b64decode(req.content)
    except Exception:
        raise HTTPException(400, "Invalid base64 content")

    async with aiofiles.open(file_path, "wb") as f:
        await f.write(raw)

    size_kb = round(os.path.getsize(file_path) / 1024, 1)

    # Delete old file for this slot
    old = await get_workbook_slot(req.slot)
    if old:
        for old_file in REFS_DIR.glob(f"{old['ref_id']}*"):
            try:
                os.unlink(old_file)
            except Exception:
                pass

    ok = await save_workbook_slot(
        slot=req.slot, ref_id=ref_id, file_name=req.name,
        file_ext=safe_ext, size_kb=size_kb, uploaded_by=req.username or "admin",
    )
    if not ok:
        raise HTTPException(500, "Failed to save workbook slot record")

    print(f"   📊 Workbook slot {req.slot} saved: {req.name} ({size_kb} KB)")
    return {"success": True, "slot": req.slot, "refId": ref_id,
            "fileName": req.name, "sizeKB": size_kb}


@app.delete("/admin/workbook-slots/{slot}")
async def route_delete_workbook_slot(slot: int):
    """Remove the grounding workbook from a specific slot."""
    existing = await get_workbook_slot(slot)
    if not existing:
        raise HTTPException(404, f"No workbook in slot {slot}")
    for old_file in REFS_DIR.glob(f"{existing['ref_id']}*"):
        try:
            os.unlink(old_file)
        except Exception:
            pass
    await delete_workbook_slot(slot)
    print(f"   🗑 Workbook slot {slot} removed")
    return {"success": True, "slot": slot}


# ═══════════════════════════════════════════════════════════════════════════════
# WORKBOOK GENERATOR ENDPOINTS
# ═══════════════════════════════════════════════════════════════════════════════

class ExtractCountriesRequest(BaseModel):
    fileData: str       # base64
    fileName: str


class GenerateWorkbookRequest(BaseModel):
    countries: list     # [{"name": "Germany", "iso3": "DEU"}, ...]
    username:  Optional[str] = "user"


@app.post("/workbook/extract-countries")
async def route_extract_countries(req: ExtractCountriesRequest):
    """Extract in-scope countries from a SOW or SP51 document using Claude AI."""
    from workbook_processor import extract_countries_from_document
    countries = await extract_countries_from_document(req.fileData, req.fileName)
    return {"success": True, "countries": countries, "count": len(countries)}


@app.get("/workbook/detect-countries")
async def route_detect_countries():
    """
    Scan all grounded reference workbooks and return the distinct countries
    actually present in their CSF sheets — used by the manual country-
    selection flow (no SOW required).
    """
    from workbook_processor import detect_countries_in_workbooks

    slots = await list_workbook_slots()
    if not slots:
        raise HTTPException(404, "No reference workbooks have been uploaded. Ask Admin to upload grounding workbooks.")

    slot_files = [
        {"slot": s["slot"], "ref_id": s["ref_id"], "file_name": s["file_name"], "file_ext": s["file_ext"]}
        for s in slots
    ]
    countries = detect_countries_in_workbooks(slot_files, REFS_DIR)
    return {"success": True, "countries": countries, "count": len(countries)}


@app.post("/workbook/generate")
async def route_generate_workbook(req: GenerateWorkbookRequest):
    """
    Generate a filtered Configuration Workbook xlsx.
    Reads all occupied workbook slots, filters CSF sheets to in-scope countries,
    returns the xlsx file as a download.
    """
    from workbook_processor import generate_filtered_workbook
    from fastapi.responses import Response

    if not req.countries:
        raise HTTPException(400, "No countries provided")

    slots = await list_workbook_slots()
    if not slots:
        raise HTTPException(404, "No reference workbooks have been uploaded. Ask Admin to upload grounding workbooks.")

    slot_files = []
    for s in slots:
        slot_files.append({
            "slot":      s["slot"],
            "ref_id":    s["ref_id"],
            "file_name": s["file_name"],
            "file_ext":  s["file_ext"],
        })

    xlsx_bytes = await generate_filtered_workbook(slot_files, req.countries, REFS_DIR)
    if not xlsx_bytes:
        raise HTTPException(500, "Workbook generation failed — check server logs")

    country_codes = "_".join(c.get("iso3", "") for c in req.countries[:5])
    file_name = f"Config_Workbook_{country_codes}.xlsx"

    print(f"   📊 Generated workbook: {file_name} ({len(xlsx_bytes)//1024} KB) for {len(req.countries)} countries")
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )
