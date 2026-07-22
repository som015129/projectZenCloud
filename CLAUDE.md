# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

ProjectZen — an AI document-generation platform for Salesforce implementation
projects. A FastAPI backend (`middleware/`) calls the Anthropic API to turn
project source documents (SOW, workshop notes, etc.) into a full set of
enterprise delivery documents (Project Plan, RTM, Solution Design, Test
Cases, Cutover Plan, ...) as real .xlsx/.docx/.pdf/.pptx/.xml files. The
frontend is a set of static, single-file HTML/JS apps with no build step.

There is no test suite, linter, or build tooling in this repo — do not
invent commands for them.

## Running the app

Backend (from `middleware/`):

```bash
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

Requires `middleware/.env` with `ANTHROPIC_API_KEY=...` (loaded manually in
`app.py` if `python-dotenv` isn't installed — see the top of `app.py`).
Health check: `GET http://localhost:8000/health`.

Frontend: plain HTML files, opened directly in a browser (no dev server,
no bundler). `frontend/index.html` is the main app — it has a "server URL"
field (defaults to `http://localhost:8000`) so it can point at any running
backend instance. Other files under `frontend/` and `project-zen/` are
alternate/prototype UIs and standalone documentation pages.

Data persistence: SQLite at `storage/documents.db` (created automatically
on startup via `init_db()` / `init_cascade_tables()` / `init_grounding_table()`
in `db.py`). Generated files land in `storage/document_store/`; uploaded
reference files in `storage/document_store/refs/`.

## Architecture

### Backend layout (`middleware/`, Tier 2 in the code's own terminology)

- `app.py` — the only FastAPI entry point; every route lives here. Reads
  `.env`, wires CORS wide open, calls `init_db()` etc. on `@app.on_event("startup")`.
- `db.py` — Tier 3: all `aiosqlite` reads/writes. One `documents` table (+
  FTS5 `documents_fts` virtual table kept in sync via triggers) for
  single-shot generations, plus addon tables for the cascade feature
  (`cascade_sessions`, `cascade_documents`, `doc_format_config`,
  `grounding_docs`, `workbook_slots`).
- `templates.py` — pure functions that turn a Claude-produced JSON "plan"
  into an actual xlsx/docx/pdf/pptx/xml file (openpyxl / python-docx /
  python-pptx / reportlab / `xml.etree`). No subprocesses.
- `extractor.py` — turns any supported input file (docx, xlsx, pptx, pdf,
  xml, txt) into plain text, both for feeding Claude and for building
  `content_summary` used in search/grounding.
- `agent.py` — the "Ask Library" feature: a real Claude tool-use loop
  (`search_library` / `read_document` / `list_recent` / `get_download_link`)
  for natural-language Q&A over the stored document library.
- `knowledge_graph.py` — a hardcoded 32-node dependency graph of every SF
  delivery document across 5 phases (Initiate → Confirm → Design-Build →
  Integrate → Deploy). Defines edges (what feeds what), per-node change
  impact/governance text, tier (T1 critical → T4 low blast radius, which
  drives Opus vs Sonnet model selection), default output format per doc
  type, and Kahn's-algorithm BFS wave computation (`compute_bfs_waves`) used
  to order parallel generation. This module is the source of truth the rest
  of the "cascade" system is built around — read it first when touching
  cascade behavior.
- `cascade_agent.py` — orchestrates multi-document "cascade" generation
  using LangGraph `StateGraph`s (one graph for full new-project generation,
  one for delta/change-propagation updates), with Anthropic calls made
  directly (not through LangGraph's model wrappers). Generates one wave
  (a set of BFS-independent documents) at a time, in parallel, streaming
  progress as SSE events. In-memory `MemorySaver` checkpointer — cascade
  session state does not survive a process restart.
- `context_store.py` — the typed `CascadeState` LangGraph state shape, plus
  a separate in-memory SSE queue registry (kept apart from LangGraph state
  because `asyncio.Queue` isn't serializable by the checkpointer).
- `delta_engine.py` — Agent 4: a Claude Opus tool-use loop that diffs an old
  vs. new document version and determines which downstream nodes (per the
  knowledge graph) are impacted, before a delta cascade run is kicked off.
- `workbook_processor.py` — SAP SuccessFactors-specific logic: extracts
  in-scope countries from a SOW via Claude, then filters up to 4
  admin-uploaded reference "Configuration Workbook" files down to a
  country-specific xlsx (CSF vs. Global sheet detection by tab-name/column
  heuristics, see `COUNTRY_ALIASES`).

### The "cascade" concept

A cascade session generates many interdependent documents from one input
document, in dependency order. Flow: `POST /cascade/start` extracts input
text → fires `run_new_cascade` in the background → client subscribes to
`GET /cascade/stream/{session_id}` (SSE) for live progress. Generation
proceeds wave-by-wave (`compute_bfs_waves`): all docs in a wave run in
parallel since their upstream dependencies are already done. A **delta**
run (`POST /cascade/delta/analyse` then `/cascade/delta/apply`) re-runs only
the subset of the graph downstream of a changed node, versioning each
affected document (v1 → v2 → ...) and asking Claude to preserve unchanged
content while only touching the diffed sections.

### Model selection

`TIER_MODELS` in `knowledge_graph.py` maps document tier → model:
T1/T2 (critical/high) → `claude-opus-4-8`, T3/T4 → `claude-sonnet-4-6`.
The single-document `/generate` route (non-cascade) always uses the
`MODEL` constant in `app.py` (currently `claude-opus-4-8`). `agent.py`'s
Ask-Library loop uses its own separate `MODEL` constant.

### CSF XML special case

`POST /generate` with `outputFormat=xml` on an Excel input whose
prompt/filename mentions CSF/SuccessFactors/HRIS keywords is routed to
`_handle_csf_xml()`, which bypasses Claude entirely and reads the CSF
sheets directly out of the workbook via `generate_csf_xml()` in
`templates.py` — deterministic, not model-generated.

### Frontend

No framework, no build step — each HTML file is self-contained (styles,
markup, and JS inline) and talks to the backend via `fetch()` against a
user-editable server URL (`srv()` helper, `#serverUrl` input,
`frontend/index.html`). `project-zen/` is a separate splash/landing UI
(`index.html` + `main.js` + `style.css`) linking out to a doc-gen page.
`frontend/documentation_*.html` and `Documentation_UserPage.html` are
static reference/help pages, not connected to the API.
