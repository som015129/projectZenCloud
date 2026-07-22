"""
workbook_processor.py — SAP SuccessFactors Configuration Workbook Processor

Handles:
  1. Country extraction from SOW / SP51 documents via Claude AI
  2. Loading grounded reference workbooks (up to 4 slots)
  3. Identifying CSF (Country-Specific Feature) vs Global sheets
  4. Filtering CSF sheets to in-scope countries only
  5. Generating a filtered Configuration Workbook as .xlsx

SAP SF Employee Central knowledge applied:
  - CSF sheets are country-specific: named with ISO codes, contain "CSF" in tab name,
    or have a Country/Country Code column as a primary key column
  - Global sheets apply to all countries (Foundation Objects, Org Structure, Pay Components, etc.)
  - Common CSF indicators: tab names like DEU, GBR, USA, SGP, AUS, or prefixes like
    EC_CSF_, Pay_Info_, Legal_Entity_, Tax_ followed by a country code
  - A "Country" or "Country Group" column is the definitive CSF row filter
"""

import io
import os
import re
import json
import base64
import struct
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import anthropic

# ── ISO country codes and common names used in SF workbooks ──────────────────
# Mapping: country name variants → ISO alpha-3 / alpha-2
COUNTRY_ALIASES: Dict[str, str] = {
    "germany": "DEU", "deutschland": "DEU", "deu": "DEU", "de": "DEU",
    "united kingdom": "GBR", "uk": "GBR", "great britain": "GBR", "gbr": "GBR", "gb": "GBR",
    "united states": "USA", "us": "USA", "usa": "USA", "united states of america": "USA",
    "france": "FRA", "fra": "FRA", "fr": "FRA",
    "netherlands": "NLD", "nld": "NLD", "nl": "NLD", "the netherlands": "NLD",
    "australia": "AUS", "aus": "AUS", "au": "AUS",
    "singapore": "SGP", "sgp": "SGP", "sg": "SGP",
    "india": "IND", "ind": "IND", "in": "IND",
    "japan": "JPN", "jpn": "JPN", "jp": "JPN",
    "canada": "CAN", "can": "CAN", "ca": "CAN",
    "china": "CHN", "chn": "CHN", "cn": "CHN", "prc": "CHN",
    "brazil": "BRA", "bra": "BRA", "br": "BRA",
    "mexico": "MEX", "mex": "MEX", "mx": "MEX",
    "spain": "ESP", "esp": "ESP", "es": "ESP",
    "italy": "ITA", "ita": "ITA", "it": "ITA",
    "sweden": "SWE", "swe": "SWE", "se": "SWE",
    "norway": "NOR", "nor": "NOR", "no": "NOR",
    "denmark": "DNK", "dnk": "DNK", "dk": "DNK",
    "finland": "FIN", "fin": "FIN", "fi": "FIN",
    "switzerland": "CHE", "che": "CHE", "ch": "CHE",
    "austria": "AUT", "aut": "AUT", "at": "AUT",
    "belgium": "BEL", "bel": "BEL", "be": "BEL",
    "poland": "POL", "pol": "POL", "pl": "POL",
    "south africa": "ZAF", "zaf": "ZAF", "za": "ZAF",
    "new zealand": "NZL", "nzl": "NZL", "nz": "NZL",
    "malaysia": "MYS", "mys": "MYS", "my": "MYS",
    "thailand": "THA", "tha": "THA", "th": "THA",
    "indonesia": "IDN", "idn": "IDN", "id": "IDN",
    "philippines": "PHL", "phl": "PHL", "ph": "PHL",
    "hong kong": "HKG", "hkg": "HKG", "hk": "HKG",
    "taiwan": "TWN", "twn": "TWN", "tw": "TWN",
    "south korea": "KOR", "kor": "KOR", "kr": "KOR", "korea": "KOR",
    "uae": "ARE", "are": "ARE", "ae": "ARE", "united arab emirates": "ARE",
    "saudi arabia": "SAU", "sau": "SAU", "sa": "SAU",
    "israel": "ISR", "isr": "ISR", "il": "ISR",
    "turkey": "TUR", "tur": "TUR", "tr": "TUR",
    "russia": "RUS", "rus": "RUS", "ru": "RUS",
    "czech republic": "CZE", "czechia": "CZE", "cze": "CZE", "cz": "CZE",
    "hungary": "HUN", "hun": "HUN", "hu": "HUN",
    "romania": "ROU", "rou": "ROU", "ro": "ROU",
    "portugal": "PRT", "prt": "PRT", "pt": "PRT",
    "greece": "GRC", "grc": "GRC", "gr": "GRC",
    "ireland": "IRL", "irl": "IRL", "ie": "IRL",
    "colombia": "COL", "col": "COL", "co": "COL",
    "argentina": "ARG", "arg": "ARG", "ar": "ARG",
    "chile": "CHL", "chl": "CHL", "cl": "CHL",
    "peru": "PER", "per": "PER", "pe": "PER",
    "egypt": "EGY", "egy": "EGY", "eg": "EGY",
    "nigeria": "NGA", "nga": "NGA", "ng": "NGA",
    "kenya": "KEN", "ken": "KEN", "ke": "KEN",
}

# Tab name patterns that definitively indicate CSF (country-specific) sheets
CSF_TAB_PATTERNS = [
    r'\bCSF\b', r'_CSF', r'CSF_',
    r'\bLegal[\s_-]?Entity\b', r'\bPay[\s_-]?Info\b',
    r'\bPayroll\b', r'\bTax[\s_-]?',
    r'\bBenefits?\b', r'\bTime[\s_-]?Off\b',
    r'\bLeave[\s_-]?',
    # ISO 3-letter codes in tab name
    r'\b(DEU|GBR|USA|FRA|NLD|AUS|SGP|IND|JPN|CAN|CHN|BRA|MEX|ESP|ITA|SWE|NOR|DNK|FIN|CHE|AUT|BEL|POL|ZAF|NZL|MYS|THA|IDN|PHL|HKG|TWN|KOR|ARE|SAU|ISR|TUR|RUS|CZE|HUN|ROU|PRT|GRC|IRL|COL|ARG|CHL|PER|EGY|NGA|KEN)\b',
]
CSF_TAB_RE = re.compile('|'.join(CSF_TAB_PATTERNS), re.IGNORECASE)

# Column header names that indicate a country filter column
COUNTRY_COL_HEADERS = {
    "country", "country code", "country group", "country/region", "countryofcompany",
    "country of company", "country_code", "country_group", "csf country",
    "legal entity country", "payroll country", "geo", "region", "locale",
    "country (iso)", "iso country", "country iso", "cc",
}

# Tab names that are always GLOBAL (never CSF-filtered)
GLOBAL_TAB_KEYWORDS = {
    "foundation", "corporate", "job_code", "job code", "job classification",
    "pay_grade", "pay grade", "pay_group", "pay group", "pay component",
    "org_chart", "org chart", "position", "department", "division", "cost_center",
    "cost center", "business_unit", "business unit", "location", "workflow",
    "dynamic_group", "dynamic group", "event_reason", "event reason",
    "employment_type", "employment type", "holiday_calendar", "holiday calendar",
    "work_schedule", "work schedule", "global", "template", "instructions",
    "readme", "change_log", "change log", "version", "legend", "overview",
    "table of contents", "toc", "index",
}


# ── File-type text extractors ────────────────────────────────────────────────

def _extract_pptx_text(raw_bytes: bytes) -> str:
    """Extract all visible text from a .pptx file (slides + notes + tables)."""
    try:
        from pptx import Presentation  # python-pptx
        from pptx.util import Inches
        prs = Presentation(io.BytesIO(raw_bytes))
        lines = []
        for slide_num, slide in enumerate(prs.slides, 1):
            lines.append(f"--- Slide {slide_num} ---")
            for shape in slide.shapes:
                # Text frames (titles, body, text boxes)
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        t = para.text.strip()
                        if t:
                            lines.append(t)
                # Tables
                if shape.has_table:
                    for row in shape.table.rows:
                        cells = [c.text.strip() for c in row.cells if c.text.strip()]
                        if cells:
                            lines.append("\t".join(cells))
                # Notes (often contain country scope in project decks)
                try:
                    if slide.has_notes_slide:
                        notes_text = slide.notes_slide.notes_text_frame.text.strip()
                        if notes_text:
                            lines.append(f"[Notes] {notes_text}")
                except Exception:
                    pass
        return "\n".join(lines)[:20000]
    except ImportError:
        return _extract_unicode_strings(raw_bytes)
    except Exception as e:
        print(f"[workbook_processor] pptx extract error: {e}")
        return _extract_unicode_strings(raw_bytes)


def _extract_ppt_text(raw_bytes: bytes) -> str:
    """
    Extract text from an old binary .ppt file.
    python-pptx does not support .ppt; we pull UTF-16 LE and ASCII strings
    directly from the binary compound document, which reliably captures
    slide text, notes, and table cells stored by PowerPoint 97-2003.
    """
    return _extract_unicode_strings(raw_bytes)


def _extract_mpp_text(raw_bytes: bytes) -> str:
    """
    Extract human-readable text from a Microsoft Project .mpp file.

    .mpp is a binary OLE Compound Document. Task names, resource names,
    notes, and custom fields are stored as UTF-16 LE strings. Countries
    appear in task names ("Germany - Deploy"), milestones, or resource pools.

    Strategy:
      1. Try python-mpxj (Java-based wrapper) for structured extraction.
      2. Fall back to Unicode/ASCII string mining if Java/mpxj not available.
    """
    # Attempt structured extraction via mpxj if available
    try:
        import mpxj  # pip install mpxj (requires Java)
        import tempfile, os as _os
        with tempfile.NamedTemporaryFile(suffix=".mpp", delete=False) as tmp:
            tmp.write(raw_bytes)
            tmp_path = tmp.name
        try:
            from mpxj import ProjectReader
            project = ProjectReader().read(tmp_path)
            lines = []
            for task in project.tasks:
                if task.name:
                    lines.append(task.name)
                if task.notes:
                    lines.append(task.notes)
            for res in project.resources:
                if res.name:
                    lines.append(res.name)
            return "\n".join(lines)[:20000]
        finally:
            _os.unlink(tmp_path)
    except Exception:
        pass

    # Fallback: mine raw Unicode strings — works well for task/milestone names
    return _extract_unicode_strings(raw_bytes)


def _extract_docx_text(raw_bytes: bytes) -> str:
    """Extract paragraph text from a .docx file via its internal XML."""
    try:
        import zipfile
        from xml.etree import ElementTree as ET
        NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        lines = []
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            for name in zf.namelist():
                if name.startswith("word/") and name.endswith(".xml"):
                    xml_bytes = zf.read(name)
                    root = ET.fromstring(xml_bytes)
                    for para in root.iter(f"{{{NS}}}p"):
                        texts = [t.text or "" for t in para.iter(f"{{{NS}}}t")]
                        line = "".join(texts).strip()
                        if line:
                            lines.append(line)
        return "\n".join(lines)[:20000]
    except Exception as e:
        print(f"[workbook_processor] docx extract error: {e}")
        return raw_bytes.decode("utf-8", errors="replace")[:15000]


def _extract_unicode_strings(raw_bytes: bytes, min_len: int = 4) -> str:
    """
    Mine printable ASCII and UTF-16 LE strings from arbitrary binary data.
    Effective for .ppt, .mpp, and other OLE compound document formats where
    country names, task names, and region labels appear as inline text.
    """
    found = set()

    # ASCII printable runs
    for m in re.finditer(rb'[ -~]{' + str(min_len).encode() + rb',}', raw_bytes):
        s = m.group().decode("ascii", errors="ignore").strip()
        if s:
            found.add(s)

    # UTF-16 LE runs (Windows-native string encoding in OLE/binary Office files)
    for m in re.finditer(rb'(?:[\x20-\x7e]\x00){' + str(min_len).encode() + rb',}', raw_bytes):
        try:
            s = m.group().decode("utf-16-le", errors="ignore").strip()
            if s:
                found.add(s)
        except Exception:
            pass

    # Sort longer strings first (task names tend to be longer than noise)
    ordered = sorted(found, key=len, reverse=True)
    return "\n".join(ordered)[:20000]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _normalise_country(raw: str) -> Optional[str]:
    """Return ISO-3 code for a country string, or None if unrecognised."""
    key = raw.strip().lower()
    return COUNTRY_ALIASES.get(key)


def _is_csf_tab_by_name(sheet_name: str) -> bool:
    """Heuristic: does this tab name look like a CSF (country-specific) sheet?"""
    low = sheet_name.lower()
    for kw in GLOBAL_TAB_KEYWORDS:
        if kw in low:
            return False
    return bool(CSF_TAB_RE.search(sheet_name))


def _find_country_column(ws) -> Optional[int]:
    """
    Scan the first 5 rows of a sheet to find the column index (1-based)
    whose header most likely represents a country/country-code field.
    Returns None if not found.
    """
    for row_idx in range(1, 6):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value and isinstance(cell.value, str):
                header = cell.value.strip().lower()
                if header in COUNTRY_COL_HEADERS:
                    return col_idx
    return None


def _get_header_row(ws) -> int:
    """Return the 1-based row index of the header row (usually 1 or 2)."""
    for r in range(1, 5):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, min(10, ws.max_column + 1))]
        non_empty = [v for v in row_vals if v is not None and str(v).strip()]
        if len(non_empty) >= 2:
            return r
    return 1


# ── Main public functions ────────────────────────────────────────────────────

async def extract_countries_from_document(file_data_b64: str, file_name: str) -> List[Dict[str, str]]:
    """
    Use Claude to extract all in-scope countries from a SOW or SP51 document.

    Returns a list of dicts: [{"name": "Germany", "iso3": "DEU"}, ...]
    """
    try:
        raw_bytes = base64.b64decode(file_data_b64)
    except Exception:
        return []

    ext = Path(file_name).suffix.lower()

    # Build the document content block(s) for Claude based on file type
    doc_content_blocks = []

    if ext in ('.xlsx', '.xls'):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
            text_parts = []
            for sheet_name in wb.sheetnames[:5]:
                ws = wb[sheet_name]
                sheet_text = f"\n[Sheet: {sheet_name}]\n"
                for row in ws.iter_rows(max_row=200, values_only=True):
                    row_str = "\t".join(str(c) if c is not None else "" for c in row)
                    if row_str.strip():
                        sheet_text += row_str + "\n"
                text_parts.append(sheet_text)
            combined = "\n".join(text_parts)[:20000]
            doc_content_blocks = [{"type": "text", "text": combined}]
        except Exception as e:
            doc_content_blocks = [{"type": "text", "text": f"[Could not parse Excel: {e}]"}]

    elif ext == '.pdf':
        # Claude natively reads PDFs — send as document block for best accuracy
        doc_content_blocks = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": file_data_b64,
                },
            }
        ]

    elif ext == '.docx':
        text = _extract_docx_text(raw_bytes)
        doc_content_blocks = [{"type": "text", "text": text}]

    elif ext == '.pptx':
        text = _extract_pptx_text(raw_bytes)
        doc_content_blocks = [{"type": "text", "text": f"[PowerPoint Presentation: {file_name}]\n\n{text}"}]

    elif ext == '.ppt':
        text = _extract_ppt_text(raw_bytes)
        doc_content_blocks = [{"type": "text", "text": f"[PowerPoint 97-2003: {file_name}]\n\n{text}"}]

    elif ext == '.mpp':
        text = _extract_mpp_text(raw_bytes)
        doc_content_blocks = [{"type": "text", "text": f"[Microsoft Project File: {file_name}]\n\n{text}"}]

    else:
        # .txt, .csv, and any other text-based format
        try:
            text = raw_bytes.decode("utf-8", errors="replace")[:20000]
        except Exception:
            text = f"[Binary file: {file_name}]"
        doc_content_blocks = [{"type": "text", "text": text}]

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

    system = (
        "You are an expert SAP SuccessFactors implementation consultant. "
        "You are reading a Statement of Work (SOW) or SP51 scope document. "
        "Your task is to extract ALL countries that are explicitly in scope for this implementation. "
        "Return ONLY a JSON array of objects with 'name' (full English country name) and 'iso3' (ISO 3166-1 alpha-3 code). "
        "Example: [{\"name\": \"Germany\", \"iso3\": \"DEU\"}, {\"name\": \"United Kingdom\", \"iso3\": \"GBR\"}]. "
        "Do not include countries that are mentioned as out-of-scope, excluded, or future phases. "
        "If the document mentions regions (e.g., 'EMEA', 'APAC'), list only the specific countries explicitly named. "
        "Return ONLY the JSON array, no other text."
    )

    content_blocks = doc_content_blocks + [
        {
            "type": "text",
            "text": (
                f"Extract all in-scope countries from this {'SOW' if 'sow' in file_name.lower() else 'SP51/scope'} document. "
                "Return a JSON array with 'name' and 'iso3' fields only."
            ),
        }
    ]

    try:
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=1024,
            system=system,
            messages=[{"role": "user", "content": content_blocks}],
        )
        raw_text = response.content[0].text.strip()
        # Strip markdown code fences if present
        raw_text = re.sub(r"^```[a-z]*\n?", "", raw_text)
        raw_text = re.sub(r"\n?```$", "", raw_text)
        countries = json.loads(raw_text)
        if isinstance(countries, list):
            # Validate and enrich
            result = []
            for c in countries:
                if isinstance(c, dict) and c.get("name"):
                    iso3 = c.get("iso3", "")
                    if not iso3:
                        iso3 = _normalise_country(c["name"]) or ""
                    result.append({"name": c["name"].strip(), "iso3": iso3.upper()})
            return result
    except Exception as e:
        print(f"[workbook_processor] country extraction error: {e}")

    return []


def load_workbook_from_path(ref_file_path: str) -> Optional[openpyxl.Workbook]:
    """Load an openpyxl workbook from disk. Returns None on failure."""
    try:
        return openpyxl.load_workbook(ref_file_path, data_only=True)
    except Exception as e:
        print(f"[workbook_processor] load error {ref_file_path}: {e}")
        return None


def classify_sheets(wb: openpyxl.Workbook) -> Dict[str, Dict]:
    """
    Classify all sheets in a workbook as:
      - 'csf': country-specific, must be filtered
      - 'global': applies to all countries, copy as-is
      - 'meta': instructions / legend / TOC, copy as-is

    Returns dict: sheet_name -> {type, country_col (if csf), header_row}
    """
    result = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        low = sheet_name.lower().strip()

        # Check for meta sheets first
        if any(kw in low for kw in ("readme", "instruction", "legend", "toc", "table of content", "version", "change log", "changelog")):
            result[sheet_name] = {"type": "meta"}
            continue

        # Check if globally excluded by name
        if any(kw in low for kw in GLOBAL_TAB_KEYWORDS):
            result[sheet_name] = {"type": "global"}
            continue

        # Try to find a country column in the sheet data
        country_col = _find_country_column(ws)
        header_row = _get_header_row(ws)

        if country_col:
            result[sheet_name] = {"type": "csf", "country_col": country_col, "header_row": header_row}
        elif _is_csf_tab_by_name(sheet_name):
            # Tab name pattern suggests CSF — scan first data column for country codes
            result[sheet_name] = {"type": "csf", "country_col": None, "header_row": header_row}
        else:
            result[sheet_name] = {"type": "global"}

    return result


def _matches_country(cell_value: Any, in_scope_iso3: set, in_scope_names: set) -> bool:
    """Check if a cell value matches one of the in-scope countries."""
    if cell_value is None:
        return False
    val = str(cell_value).strip().lower()
    if not val:
        return False
    # Direct ISO-3 match
    if val.upper() in in_scope_iso3:
        return True
    # Full name match
    if val in in_scope_names:
        return True
    # Alias lookup
    mapped = _normalise_country(val)
    if mapped and mapped in in_scope_iso3:
        return True
    return False


def filter_workbook(
    source_wb: openpyxl.Workbook,
    in_scope_countries: List[Dict[str, str]],
    label: str = "Workbook",
) -> openpyxl.Workbook:
    """
    Create a new workbook containing:
    - All global sheets copied as-is
    - All meta sheets copied as-is
    - CSF sheets filtered to only rows whose country column is in in_scope_countries

    Returns a new openpyxl.Workbook.
    """
    iso3_set = {c.get("iso3", "").upper() for c in in_scope_countries if c.get("iso3")}
    name_set  = {c.get("name", "").lower() for c in in_scope_countries if c.get("name")}

    sheet_classification = classify_sheets(source_wb)
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)  # remove default blank sheet

    # Copy each sheet
    for sheet_name in source_wb.sheetnames:
        info = sheet_classification.get(sheet_name, {"type": "global"})
        src_ws = source_wb[sheet_name]
        dst_ws = out_wb.create_sheet(title=sheet_name)

        if info["type"] in ("global", "meta"):
            # Copy everything
            for row in src_ws.iter_rows():
                for cell in row:
                    dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        try:
                            dst_cell.font      = cell.font.copy()
                            dst_cell.fill      = cell.fill.copy()
                            dst_cell.border    = cell.border.copy()
                            dst_cell.alignment = cell.alignment.copy()
                            dst_cell.number_format = cell.number_format
                        except Exception:
                            pass
        else:
            # CSF sheet — filter rows
            country_col  = info.get("country_col")
            header_row   = info.get("header_row", 1)
            out_row_idx  = 0
            kept_rows    = 0

            for src_row in src_ws.iter_rows():
                row_num = src_row[0].row if src_row else 0

                # Always copy rows up to and including header
                if row_num <= header_row:
                    out_row_idx += 1
                    for cell in src_row:
                        dst_cell = dst_ws.cell(row=out_row_idx, column=cell.column, value=cell.value)
                        if cell.has_style:
                            try:
                                dst_cell.font      = cell.font.copy()
                                dst_cell.fill      = cell.fill.copy()
                                dst_cell.border    = cell.border.copy()
                                dst_cell.alignment = cell.alignment.copy()
                                dst_cell.number_format = cell.number_format
                            except Exception:
                                pass
                    continue

                # For data rows: check country filter
                include_row = False
                if country_col:
                    cv = src_row[country_col - 1].value if len(src_row) >= country_col else None
                    include_row = _matches_country(cv, iso3_set, name_set)
                else:
                    # No explicit country column found — check all cells in the row
                    # for a value matching an in-scope country (conservative: keep if any match)
                    for cell in src_row[:5]:
                        if _matches_country(cell.value, iso3_set, name_set):
                            include_row = True
                            break

                if include_row:
                    out_row_idx += 1
                    kept_rows += 1
                    for cell in src_row:
                        dst_cell = dst_ws.cell(row=out_row_idx, column=cell.column, value=cell.value)
                        if cell.has_style:
                            try:
                                dst_cell.font      = cell.font.copy()
                                dst_cell.fill      = cell.fill.copy()
                                dst_cell.border    = cell.border.copy()
                                dst_cell.alignment = cell.alignment.copy()
                                dst_cell.number_format = cell.number_format
                            except Exception:
                                pass

            print(f"  [CSF filter] {sheet_name}: kept {kept_rows} rows for {len(iso3_set)} countries")

        # Copy column widths
        try:
            for col_letter, col_dim in src_ws.column_dimensions.items():
                dst_ws.column_dimensions[col_letter].width = col_dim.width
        except Exception:
            pass

    return out_wb


async def generate_filtered_workbook(
    slot_files: List[Dict],  # [{"slot": 1, "path": "...", "file_name": "...", "ref_id": "..."}, ...]
    in_scope_countries: List[Dict[str, str]],
    refs_dir: Path,
) -> Optional[bytes]:
    """
    Load each slot file, filter CSF sheets to in-scope countries,
    and produce a combined .xlsx output as bytes.

    If only one slot, output is the filtered workbook directly.
    If multiple slots, create a combined workbook with a summary sheet first,
    then all filtered sheets from each slot (prefixed with slot number if name collides).

    Returns raw xlsx bytes or None on failure.
    """
    if not slot_files or not in_scope_countries:
        return None

    country_labels = ", ".join(c.get("iso3", c.get("name", "")) for c in in_scope_countries)
    print(f"[workbook_processor] Generating for countries: {country_labels}")

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    # ── Summary sheet ────────────────────────────────────────────────────────
    summary = out_wb.create_sheet("_Summary")
    summary["A1"] = "Configuration Workbook — Country-Filtered Output"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A3"] = "In-Scope Countries:"
    summary["A3"].font = Font(bold=True)
    for i, c in enumerate(in_scope_countries, start=4):
        summary[f"A{i}"] = c.get("name", "")
        summary[f"B{i}"] = c.get("iso3", "")
    summary["A3"].font = Font(bold=True)
    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 10

    import datetime as dt_mod
    summary[f"A{len(in_scope_countries)+5}"] = f"Generated: {dt_mod.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
    summary[f"A{len(in_scope_countries)+6}"] = f"Source workbooks: {len(slot_files)}"

    existing_sheet_names: set = {"_Summary"}

    for slot_info in slot_files:
        slot_num  = slot_info.get("slot", 1)
        file_name = slot_info.get("file_name", f"Workbook_{slot_num}")
        ref_id    = slot_info.get("ref_id", "")
        file_ext  = slot_info.get("file_ext", "xlsx")

        # Find file on disk
        file_path = None
        for candidate in refs_dir.glob(f"{ref_id}.*"):
            file_path = candidate
            break

        if not file_path or not file_path.exists():
            print(f"  [slot {slot_num}] file not found: {ref_id}")
            continue

        src_wb = load_workbook_from_path(str(file_path))
        if not src_wb:
            continue

        filtered_wb = filter_workbook(src_wb, in_scope_countries, label=f"Slot {slot_num}")

        # Copy sheets into combined workbook
        for sheet_name in filtered_wb.sheetnames:
            # Handle name collision when multiple slots have same sheet names
            dest_name = sheet_name
            if len(slot_files) > 1:
                prefix = f"S{slot_num}_"
                dest_name = (prefix + sheet_name)[:31]
            # Ensure uniqueness
            if dest_name in existing_sheet_names:
                dest_name = (dest_name[:28] + f"_{slot_num}")[:31]
            existing_sheet_names.add(dest_name)

            src_ws  = filtered_wb[sheet_name]
            dst_ws  = out_wb.create_sheet(title=dest_name)

            for row in src_ws.iter_rows():
                for cell in row:
                    dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        try:
                            dst_cell.font      = cell.font.copy()
                            dst_cell.fill      = cell.fill.copy()
                            dst_cell.border    = cell.border.copy()
                            dst_cell.alignment = cell.alignment.copy()
                            dst_cell.number_format = cell.number_format
                        except Exception:
                            pass
            try:
                for col_letter, col_dim in src_ws.column_dimensions.items():
                    dst_ws.column_dimensions[col_letter].width = col_dim.width
            except Exception:
                pass

    if len(out_wb.sheetnames) <= 1:
        # Only summary sheet — no workbook data was filtered
        print("[workbook_processor] No sheets produced after filtering")

    buf = io.BytesIO()
    out_wb.save(buf)
    return buf.getvalue()
