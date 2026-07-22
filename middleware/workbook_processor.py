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
import asyncio
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import anthropic

# ── ISO 3166-1 country reference (alpha-2, alpha-3, primary name) ────────────
# Full ISO country list — SAP SF Employee Central ships 100+ localizations, so
# detection must not be limited to a hand-picked subset (that previously
# caused most in-scope countries to be silently dropped during workbook scan).
_ISO_COUNTRIES: List[Tuple[str, str, str]] = [
    ("AF", "AFG", "Afghanistan"), ("AL", "ALB", "Albania"), ("DZ", "DZA", "Algeria"),
    ("AS", "ASM", "American Samoa"), ("AD", "AND", "Andorra"), ("AO", "AGO", "Angola"),
    ("AI", "AIA", "Anguilla"), ("AG", "ATG", "Antigua and Barbuda"), ("AR", "ARG", "Argentina"),
    ("AM", "ARM", "Armenia"), ("AW", "ABW", "Aruba"), ("AU", "AUS", "Australia"),
    ("AT", "AUT", "Austria"), ("AZ", "AZE", "Azerbaijan"), ("BS", "BHS", "Bahamas"),
    ("BH", "BHR", "Bahrain"), ("BD", "BGD", "Bangladesh"), ("BB", "BRB", "Barbados"),
    ("BY", "BLR", "Belarus"), ("BE", "BEL", "Belgium"), ("BZ", "BLZ", "Belize"),
    ("BJ", "BEN", "Benin"), ("BM", "BMU", "Bermuda"), ("BT", "BTN", "Bhutan"),
    ("BO", "BOL", "Bolivia"), ("BA", "BIH", "Bosnia and Herzegovina"), ("BW", "BWA", "Botswana"),
    ("BR", "BRA", "Brazil"), ("BN", "BRN", "Brunei"), ("BG", "BGR", "Bulgaria"),
    ("BF", "BFA", "Burkina Faso"), ("BI", "BDI", "Burundi"), ("KH", "KHM", "Cambodia"),
    ("CM", "CMR", "Cameroon"), ("CA", "CAN", "Canada"), ("CV", "CPV", "Cape Verde"),
    ("KY", "CYM", "Cayman Islands"), ("CF", "CAF", "Central African Republic"), ("TD", "TCD", "Chad"),
    ("CL", "CHL", "Chile"), ("CN", "CHN", "China"), ("CO", "COL", "Colombia"),
    ("KM", "COM", "Comoros"), ("CG", "COG", "Congo"), ("CD", "COD", "Congo, Democratic Republic of the"),
    ("CR", "CRI", "Costa Rica"), ("CI", "CIV", "Cote d'Ivoire"), ("HR", "HRV", "Croatia"),
    ("CU", "CUB", "Cuba"), ("CW", "CUW", "Curacao"), ("CY", "CYP", "Cyprus"),
    ("CZ", "CZE", "Czech Republic"), ("DK", "DNK", "Denmark"), ("DJ", "DJI", "Djibouti"),
    ("DM", "DMA", "Dominica"), ("DO", "DOM", "Dominican Republic"), ("EC", "ECU", "Ecuador"),
    ("EG", "EGY", "Egypt"), ("SV", "SLV", "El Salvador"), ("GQ", "GNQ", "Equatorial Guinea"),
    ("ER", "ERI", "Eritrea"), ("EE", "EST", "Estonia"), ("SZ", "SWZ", "Eswatini"),
    ("ET", "ETH", "Ethiopia"), ("FJ", "FJI", "Fiji"), ("FI", "FIN", "Finland"),
    ("FR", "FRA", "France"), ("GA", "GAB", "Gabon"), ("GM", "GMB", "Gambia"),
    ("GE", "GEO", "Georgia"), ("DE", "DEU", "Germany"), ("GH", "GHA", "Ghana"),
    ("GI", "GIB", "Gibraltar"), ("GR", "GRC", "Greece"), ("GL", "GRL", "Greenland"),
    ("GD", "GRD", "Grenada"), ("GU", "GUM", "Guam"), ("GT", "GTM", "Guatemala"),
    ("GG", "GGY", "Guernsey"), ("GN", "GIN", "Guinea"), ("GW", "GNB", "Guinea-Bissau"),
    ("GY", "GUY", "Guyana"), ("HT", "HTI", "Haiti"), ("HN", "HND", "Honduras"),
    ("HK", "HKG", "Hong Kong"), ("HU", "HUN", "Hungary"), ("IS", "ISL", "Iceland"),
    ("IN", "IND", "India"), ("ID", "IDN", "Indonesia"), ("IR", "IRN", "Iran"),
    ("IQ", "IRQ", "Iraq"), ("IE", "IRL", "Ireland"), ("IM", "IMN", "Isle of Man"),
    ("IL", "ISR", "Israel"), ("IT", "ITA", "Italy"), ("JM", "JAM", "Jamaica"),
    ("JP", "JPN", "Japan"), ("JE", "JEY", "Jersey"), ("JO", "JOR", "Jordan"),
    ("KZ", "KAZ", "Kazakhstan"), ("KE", "KEN", "Kenya"), ("KI", "KIR", "Kiribati"),
    ("KP", "PRK", "North Korea"), ("KR", "KOR", "South Korea"), ("KW", "KWT", "Kuwait"),
    ("KG", "KGZ", "Kyrgyzstan"), ("LA", "LAO", "Laos"), ("LV", "LVA", "Latvia"),
    ("LB", "LBN", "Lebanon"), ("LS", "LSO", "Lesotho"), ("LR", "LBR", "Liberia"),
    ("LY", "LBY", "Libya"), ("LI", "LIE", "Liechtenstein"), ("LT", "LTU", "Lithuania"),
    ("LU", "LUX", "Luxembourg"), ("MO", "MAC", "Macao"), ("MG", "MDG", "Madagascar"),
    ("MW", "MWI", "Malawi"), ("MY", "MYS", "Malaysia"), ("MV", "MDV", "Maldives"),
    ("ML", "MLI", "Mali"), ("MT", "MLT", "Malta"), ("MH", "MHL", "Marshall Islands"),
    ("MR", "MRT", "Mauritania"), ("MU", "MUS", "Mauritius"), ("MX", "MEX", "Mexico"),
    ("FM", "FSM", "Micronesia"), ("MD", "MDA", "Moldova"), ("MC", "MCO", "Monaco"),
    ("MN", "MNG", "Mongolia"), ("ME", "MNE", "Montenegro"), ("MA", "MAR", "Morocco"),
    ("MZ", "MOZ", "Mozambique"), ("MM", "MMR", "Myanmar"), ("NA", "NAM", "Namibia"),
    ("NR", "NRU", "Nauru"), ("NP", "NPL", "Nepal"), ("NL", "NLD", "Netherlands"),
    ("NC", "NCL", "New Caledonia"), ("NZ", "NZL", "New Zealand"), ("NI", "NIC", "Nicaragua"),
    ("NE", "NER", "Niger"), ("NG", "NGA", "Nigeria"), ("MK", "MKD", "North Macedonia"),
    ("NO", "NOR", "Norway"), ("OM", "OMN", "Oman"), ("PK", "PAK", "Pakistan"),
    ("PW", "PLW", "Palau"), ("PS", "PSE", "Palestine"), ("PA", "PAN", "Panama"),
    ("PG", "PNG", "Papua New Guinea"), ("PY", "PRY", "Paraguay"), ("PE", "PER", "Peru"),
    ("PH", "PHL", "Philippines"), ("PL", "POL", "Poland"), ("PT", "PRT", "Portugal"),
    ("PR", "PRI", "Puerto Rico"), ("QA", "QAT", "Qatar"), ("RO", "ROU", "Romania"),
    ("RU", "RUS", "Russia"), ("RW", "RWA", "Rwanda"), ("KN", "KNA", "Saint Kitts and Nevis"),
    ("LC", "LCA", "Saint Lucia"), ("VC", "VCT", "Saint Vincent and the Grenadines"),
    ("WS", "WSM", "Samoa"), ("SM", "SMR", "San Marino"), ("ST", "STP", "Sao Tome and Principe"),
    ("SA", "SAU", "Saudi Arabia"), ("SN", "SEN", "Senegal"), ("RS", "SRB", "Serbia"),
    ("SC", "SYC", "Seychelles"), ("SL", "SLE", "Sierra Leone"), ("SG", "SGP", "Singapore"),
    ("SK", "SVK", "Slovakia"), ("SI", "SVN", "Slovenia"), ("SB", "SLB", "Solomon Islands"),
    ("SO", "SOM", "Somalia"), ("ZA", "ZAF", "South Africa"), ("SS", "SSD", "South Sudan"),
    ("ES", "ESP", "Spain"), ("LK", "LKA", "Sri Lanka"), ("SD", "SDN", "Sudan"),
    ("SR", "SUR", "Suriname"), ("SE", "SWE", "Sweden"), ("CH", "CHE", "Switzerland"),
    ("SY", "SYR", "Syria"), ("TW", "TWN", "Taiwan"), ("TJ", "TJK", "Tajikistan"),
    ("TZ", "TZA", "Tanzania"), ("TH", "THA", "Thailand"), ("TL", "TLS", "Timor-Leste"),
    ("TG", "TGO", "Togo"), ("TO", "TON", "Tonga"), ("TT", "TTO", "Trinidad and Tobago"),
    ("TN", "TUN", "Tunisia"), ("TR", "TUR", "Turkey"), ("TM", "TKM", "Turkmenistan"),
    ("TV", "TUV", "Tuvalu"), ("UG", "UGA", "Uganda"), ("UA", "UKR", "Ukraine"),
    ("AE", "ARE", "United Arab Emirates"), ("GB", "GBR", "United Kingdom"), ("US", "USA", "United States"),
    ("UY", "URY", "Uruguay"), ("UZ", "UZB", "Uzbekistan"), ("VU", "VUT", "Vanuatu"),
    ("VA", "VAT", "Vatican City"), ("VE", "VEN", "Venezuela"), ("VN", "VNM", "Vietnam"),
    ("VG", "VGB", "British Virgin Islands"), ("VI", "VIR", "U.S. Virgin Islands"), ("YE", "YEM", "Yemen"),
    ("ZM", "ZMB", "Zambia"), ("ZW", "ZWE", "Zimbabwe"),
]

# Colloquial / historical name variants not covered by the primary ISO name above.
_EXTRA_COUNTRY_ALIASES: Dict[str, str] = {
    "uk": "GBR", "great britain": "GBR", "britain": "GBR",
    "united states of america": "USA", "america": "USA",
    "prc": "CHN", "mainland china": "CHN",
    "korea": "KOR", "republic of korea": "KOR", "dprk": "PRK",
    "the netherlands": "NLD", "holland": "NLD",
    "ivory coast": "CIV", "czechia": "CZE", "swaziland": "SWZ", "burma": "MMR",
    "dr congo": "COD", "drc": "COD", "zaire": "COD", "republic of the congo": "COG",
    "macedonia": "MKD", "fyrom": "MKD", "vatican": "VAT", "holy see": "VAT",
    "russian federation": "RUS", "syrian arab republic": "SYR", "lao pdr": "LAO",
    "brunei darussalam": "BRN", "viet nam": "VNM", "east timor": "TLS",
    "palestinian territories": "PSE",
}


def _build_country_aliases() -> Dict[str, str]:
    aliases: Dict[str, str] = {}
    for alpha2, alpha3, name in _ISO_COUNTRIES:
        aliases[name.lower()] = alpha3
        aliases[alpha2.lower()] = alpha3
        aliases[alpha3.lower()] = alpha3
    aliases.update(_EXTRA_COUNTRY_ALIASES)
    return aliases


# Mapping: country name variants → ISO alpha-3 / alpha-2
COUNTRY_ALIASES: Dict[str, str] = _build_country_aliases()

# All ISO alpha-3 codes as a single alternation, for spotting an embedded
# country code in a sheet tab name (e.g. "DEU_Payroll") — shared by
# CSF_TAB_RE below and the tab-name fallback in detect_countries_in_workbooks.
_ISO3_ALTERNATION = '|'.join(sorted({alpha3 for _, alpha3, _ in _ISO_COUNTRIES}))
ISO3_TAB_RE = re.compile(r'\b(' + _ISO3_ALTERNATION + r')\b')

# Tab name patterns that definitively indicate CSF (country-specific) sheets
CSF_TAB_PATTERNS = [
    r'\bCSF\b', r'_CSF', r'CSF_',
    r'\bLegal[\s_-]?Entity\b', r'\bPay[\s_-]?Info\b',
    r'\bPayroll\b', r'\bTax[\s_-]?',
    r'\bBenefits?\b', r'\bTime[\s_-]?Off\b',
    r'\bLeave[\s_-]?',
    # ISO 3-letter codes in tab name
    r'\b(?:' + _ISO3_ALTERNATION + r')\b',
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


# Canonical display name per ISO-3 code (for countries detected directly from
# workbook data, where we only have a code/alias and need something to show).
# Derived from the same master list COUNTRY_ALIASES is built from, so every
# code that can be *detected* also has a display name.
ISO3_TO_NAME: Dict[str, str] = {alpha3: name for _, alpha3, name in _ISO_COUNTRIES}


# ── Helpers ───────────────────────────────────────────────────────────────────

_NAME_CODE_RE = re.compile(r'^(.*?)\s*\(([A-Za-z]{2,3})\)\s*$')


def _normalise_country(raw: str) -> Optional[str]:
    """
    Return ISO-3 code for a country string, or None if unrecognised.

    Real SF EC workbooks' CSF country columns commonly format values as
    "Angola (AGO)", "Argentina (ARG)" — name plus code in parentheses,
    not a bare name or code (verified directly against the production
    "National ID" sheet). Without this, exact-match lookup fails for
    every such value, silently excluding all CSF data regardless of
    which countries are actually in scope.
    """
    key = raw.strip().lower()
    direct = COUNTRY_ALIASES.get(key)
    if direct:
        return direct

    m = _NAME_CODE_RE.match(raw.strip())
    if m:
        name_part, code_part = m.group(1).strip().lower(), m.group(2).strip().lower()
        return COUNTRY_ALIASES.get(code_part) or COUNTRY_ALIASES.get(name_part)
    return None


def _resolve_iso3(value: Any) -> Optional[str]:
    """Resolve a raw cell value (name, alias, or ISO code) to ISO-3, or None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    return _normalise_country(s)


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

    Uses bounded iter_rows() rather than ws.cell(row=, column=) random
    access — the latter isn't supported on read-only worksheets, which
    is how sheets are loaded (see load_workbook_from_path): eager-mode
    loading of a real SF EC workbook was measured at 250-400s, versus
    15-30s in read-only mode.
    """
    for row in ws.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                header = cell.value.strip().lower()
                if header in COUNTRY_COL_HEADERS:
                    return cell.column
    return None


def _get_header_row(ws) -> int:
    """
    Return the 1-based row index of the header row (usually 1 or 2).

    Row index comes from enumerate(), not any individual cell's .row —
    a row's first element can be a lightweight EmptyCell placeholder
    (read-only mode's stand-in for a blank position within a row's
    populated range), which doesn't carry a .row attribute.
    """
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=4), start=1):
        non_empty = [c.value for c in row[:10] if c.value is not None and str(c.value).strip()]
        if len(non_empty) >= 2:
            return row_idx
    return 1


def _stream_real_rows(ws, max_gap: int = 10000):
    """
    Yield (row_idx, [(col_idx, cell), ...]) only for rows that actually
    contain at least one populated (value or styled) cell, reading the
    sheet forward and stopping early once `max_gap` consecutive fully-
    empty rows have been seen.

    Real-world Excel workbooks routinely have a handful of cells stranded
    far outside their real content block — e.g. one leftover cell near
    the sheet's absolute row limit, which genuinely exists but drags
    ws.max_row out past a million, even though real data ends a few
    dozen or few thousand rows in. Naively iterating to ws.max_row would
    mean streaming through — or, in eager/non-read-only mode, having
    openpyxl backfill a dense grid of Cell objects for — over a million
    phantom rows. The early-break-on-gap approach avoids ever reaching
    that phantom range at all, for both read-only (cheap either way,
    just avoids wasted iterations) and eager worksheets (avoids the
    expensive backfill entirely).

    Verified against a real SF EC workbook: a legitimately dense
    40,008-row sheet has zero gaps over 50 rows (so is read in full),
    while a sheet with real content ending at row ~1,062 had a single
    stray cell at row 1,048,573 — a >1,000,000-row gap that's cut here.
    """
    empty_streak = 0
    start_row = getattr(ws, "min_row", None) or 1
    for row_idx, row in enumerate(ws.iter_rows(), start=start_row):
        # Read-only mode yields lightweight EmptyCell placeholders for
        # blank positions within a row's populated range — they lack
        # .has_style (and .row) entirely (only ReadOnlyCell/Cell have
        # them), which is why row_idx comes from enumerate() above, not
        # from any individual cell's .row attribute.
        populated = [(c.column, c) for c in row if c.value is not None or getattr(c, "has_style", False)]
        if populated:
            empty_streak = 0
            yield row_idx, populated
        else:
            empty_streak += 1
            if empty_streak > max_gap:
                return


def _detached_style(value):
    """
    Return a detached, independently-assignable copy of a Font/Fill/
    Border/Alignment style value.

    On an eager-mode cell, cell.font (etc.) returns a live StyleProxy
    tied to that cell's position in the shared style table — copy() is
    required to get a plain, independently assignable object. On a
    read-only cell, cell.font already returns a plain, detached object
    with no .copy() method (AttributeError), because there's no live
    style table position to be proxying in the first place. Handle both
    by copying only when the object actually supports it.
    """
    copy_fn = getattr(value, "copy", None)
    return copy_fn() if copy_fn else value


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
    """
    Load an openpyxl workbook from disk. Returns None on failure.

    read_only=True: eager (writable) loading was measured at 250-400s
    against a real ~3.3MB, 32-sheet SF EC workbook, versus 15-30s in
    read-only mode — eager mode builds a full mutable Cell object graph
    with resolved styles for every touched position, which is far more
    expensive than read-only's streaming SAX-based access. Read-only
    cells still expose .value, .font, .fill, .border, .alignment, and
    .has_style (verified directly against this workbook), so nothing
    downstream that only reads cells is affected — only random access
    via ws.cell(row=, column=) is unavailable, which is why
    _find_country_column/_get_header_row/_stream_real_rows all use
    iter_rows() instead.
    """
    try:
        return openpyxl.load_workbook(ref_file_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"[workbook_processor] load error {ref_file_path}: {e}")
        return None


def detect_countries_in_workbooks(slot_files: List[Dict], refs_dir: Path) -> List[Dict[str, Any]]:
    """
    Scan all grounded reference workbooks and return the distinct countries
    actually present in their CSF sheets — for the manual country-selection
    flow, where no SOW is uploaded and countries must be read from the data
    that is already grounded.

    Detection order per CSF sheet:
      1. If a country column was found, resolve every data-row value in it.
      2. Otherwise, the country is usually encoded in the tab name itself
         (e.g. "DEU_Payroll") — check for an embedded ISO-3 code there.
      3. As a last resort, scan the first few cells of each data row.

    Returns: [{"name": "Germany", "iso3": "DEU", "count": 42}, ...]
             sorted alphabetically by name.
    """
    found: Dict[str, int] = {}

    for slot_info in slot_files:
        ref_id = slot_info.get("ref_id", "")
        file_path = None
        for candidate in refs_dir.glob(f"{ref_id}.*"):
            file_path = candidate
            break
        if not file_path or not file_path.exists():
            continue

        wb = load_workbook_from_path(str(file_path))
        if not wb:
            continue

        classification = classify_sheets(wb)
        for sheet_name, info in classification.items():
            if info["type"] != "csf":
                continue
            ws = wb[sheet_name]
            country_col = info.get("country_col")
            header_row  = info.get("header_row", 1)

            if country_col:
                for row_idx, col_cells in _stream_real_rows(ws):
                    if row_idx <= header_row:
                        continue
                    for col_idx, cell in col_cells:
                        if col_idx == country_col:
                            iso3 = _resolve_iso3(cell.value)
                            if iso3:
                                found[iso3] = found.get(iso3, 0) + 1
                            break
            else:
                tab_match = ISO3_TAB_RE.search(sheet_name.upper())
                if tab_match:
                    iso3 = tab_match.group(1)
                    found[iso3] = found.get(iso3, 0) + 1
                else:
                    for row_idx, col_cells in _stream_real_rows(ws):
                        if row_idx <= header_row:
                            continue
                        for col_idx, cell in col_cells:
                            if col_idx > 5:
                                break
                            iso3 = _resolve_iso3(cell.value)
                            if iso3:
                                found[iso3] = found.get(iso3, 0) + 1

    result = [
        {"name": ISO3_TO_NAME.get(iso3, iso3), "iso3": iso3, "count": count}
        for iso3, count in found.items()
    ]
    result.sort(key=lambda c: c["name"])
    return result


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
            # Copy only genuinely populated cells — not a dense
            # ws.max_row x ws.max_column grid, which can be orders of
            # magnitude larger than actual content (see _stream_real_rows).
            for row_idx, col_cells in _stream_real_rows(src_ws):
                for col_idx, cell in col_cells:
                    dst_cell = dst_ws.cell(row=row_idx, column=col_idx, value=cell.value)
                    if cell.has_style:
                        try:
                            dst_cell.font      = _detached_style(cell.font)
                            dst_cell.fill      = _detached_style(cell.fill)
                            dst_cell.border    = _detached_style(cell.border)
                            dst_cell.alignment = _detached_style(cell.alignment)
                            dst_cell.number_format = cell.number_format
                        except Exception:
                            pass
        else:
            # CSF sheet — filter rows
            country_col  = info.get("country_col")
            header_row   = info.get("header_row", 1)
            out_row_idx  = 0
            kept_rows    = 0

            for row_idx, col_cells in _stream_real_rows(src_ws):
                # Always copy rows up to and including header
                if row_idx <= header_row:
                    out_row_idx += 1
                    for col_idx, cell in col_cells:
                        dst_cell = dst_ws.cell(row=out_row_idx, column=col_idx, value=cell.value)
                        if cell.has_style:
                            try:
                                dst_cell.font      = _detached_style(cell.font)
                                dst_cell.fill      = _detached_style(cell.fill)
                                dst_cell.border    = _detached_style(cell.border)
                                dst_cell.alignment = _detached_style(cell.alignment)
                                dst_cell.number_format = cell.number_format
                            except Exception:
                                pass
                    continue

                # For data rows: check country filter
                include_row = False
                if country_col:
                    cv = None
                    for col_idx, cell in col_cells:
                        if col_idx == country_col:
                            cv = cell.value
                            break
                    include_row = _matches_country(cv, iso3_set, name_set)
                else:
                    # No explicit country column found — check the first 5
                    # columns of the row for a value matching an in-scope
                    # country (conservative: keep if any match)
                    for col_idx, cell in col_cells:
                        if col_idx > 5:
                            break
                        if _matches_country(cell.value, iso3_set, name_set):
                            include_row = True
                            break

                if include_row:
                    out_row_idx += 1
                    kept_rows += 1
                    for col_idx, cell in col_cells:
                        dst_cell = dst_ws.cell(row=out_row_idx, column=col_idx, value=cell.value)
                        if cell.has_style:
                            try:
                                dst_cell.font      = _detached_style(cell.font)
                                dst_cell.fill      = _detached_style(cell.fill)
                                dst_cell.border    = _detached_style(cell.border)
                                dst_cell.alignment = _detached_style(cell.alignment)
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


def _add_summary_sheet(out_wb: openpyxl.Workbook, in_scope_countries: List[Dict[str, str]], source_count: int) -> None:
    summary = out_wb.create_sheet("_Summary", 0)
    summary["A1"] = "Configuration Workbook — Country-Filtered Output"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A3"] = "In-Scope Countries:"
    summary["A3"].font = Font(bold=True)
    for i, c in enumerate(in_scope_countries, start=4):
        summary[f"A{i}"] = c.get("name", "")
        summary[f"B{i}"] = c.get("iso3", "")
    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 10

    import datetime as dt_mod
    summary[f"A{len(in_scope_countries)+5}"] = f"Generated: {dt_mod.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
    summary[f"A{len(in_scope_countries)+6}"] = f"Source workbooks: {source_count}"


def _find_slot_file(slot_info: Dict, refs_dir: Path) -> Optional[Path]:
    ref_id = slot_info.get("ref_id", "")
    for candidate in refs_dir.glob(f"{ref_id}.*"):
        return candidate
    return None


def _generate_filtered_workbook_sync(
    slot_files: List[Dict],
    in_scope_countries: List[Dict[str, str]],
    refs_dir: Path,
) -> Optional[bytes]:
    """
    Load each slot file, filter CSF sheets to in-scope countries,
    and produce a combined .xlsx output as bytes.

    If only one slot, the filtered workbook is used directly (a summary sheet
    is inserted at the front) — no second cell-by-cell copy pass, since that
    was doubling the (already expensive, style-preserving) copy work for the
    common case and could stall the request on large real-world workbooks.
    If multiple slots, sheets from each are merged into one combined workbook
    (prefixed with slot number if name collides).

    Returns raw xlsx bytes or None on failure.
    """
    country_labels = ", ".join(c.get("iso3", c.get("name", "")) for c in in_scope_countries)
    print(f"[workbook_processor] Generating for countries: {country_labels}")

    # ── Single slot: filter in place, skip the redundant second copy ──────────
    if len(slot_files) == 1:
        slot_info = slot_files[0]
        file_path = _find_slot_file(slot_info, refs_dir)
        if not file_path or not file_path.exists():
            print(f"  [slot {slot_info.get('slot', 1)}] file not found: {slot_info.get('ref_id', '')}")
            return None

        src_wb = load_workbook_from_path(str(file_path))
        if not src_wb:
            return None

        out_wb = filter_workbook(src_wb, in_scope_countries, label=f"Slot {slot_info.get('slot', 1)}")
        _add_summary_sheet(out_wb, in_scope_countries, source_count=1)

        buf = io.BytesIO()
        out_wb.save(buf)
        return buf.getvalue()

    # ── Multiple slots: merge filtered sheets from each into one workbook ─────
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    _add_summary_sheet(out_wb, in_scope_countries, source_count=len(slot_files))
    existing_sheet_names: set = {"_Summary"}

    for slot_info in slot_files:
        slot_num  = slot_info.get("slot", 1)
        file_path = _find_slot_file(slot_info, refs_dir)

        if not file_path or not file_path.exists():
            print(f"  [slot {slot_num}] file not found: {slot_info.get('ref_id', '')}")
            continue

        src_wb = load_workbook_from_path(str(file_path))
        if not src_wb:
            continue

        filtered_wb = filter_workbook(src_wb, in_scope_countries, label=f"Slot {slot_num}")

        # Copy sheets into combined workbook
        for sheet_name in filtered_wb.sheetnames:
            # Handle name collision when multiple slots have same sheet names
            dest_name = f"S{slot_num}_{sheet_name}"[:31]
            if dest_name in existing_sheet_names:
                dest_name = (dest_name[:28] + f"_{slot_num}")[:31]
            existing_sheet_names.add(dest_name)

            src_ws  = filtered_wb[sheet_name]
            dst_ws  = out_wb.create_sheet(title=dest_name)

            for row_idx, col_cells in _stream_real_rows(src_ws):
                for col_idx, cell in col_cells:
                    dst_cell = dst_ws.cell(row=row_idx, column=col_idx, value=cell.value)
                    if cell.has_style:
                        try:
                            dst_cell.font      = _detached_style(cell.font)
                            dst_cell.fill      = _detached_style(cell.fill)
                            dst_cell.border    = _detached_style(cell.border)
                            dst_cell.alignment = _detached_style(cell.alignment)
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


async def generate_filtered_workbook(
    slot_files: List[Dict],  # [{"slot": 1, "path": "...", "file_name": "...", "ref_id": "..."}, ...]
    in_scope_countries: List[Dict[str, str]],
    refs_dir: Path,
) -> Optional[bytes]:
    """
    Async wrapper around _generate_filtered_workbook_sync.

    The actual filtering/copying is synchronous, CPU-bound openpyxl work that
    can take a long time on real-world multi-sheet, multi-thousand-row SF EC
    workbooks. Running it inline in the event loop would block the single
    uvicorn worker for the whole duration — starving health checks and every
    other in-flight request, and risking the connection being dropped before
    a response is ever sent. asyncio.to_thread moves it off the event loop.
    """
    if not slot_files or not in_scope_countries:
        return None
    return await asyncio.to_thread(_generate_filtered_workbook_sync, slot_files, in_scope_countries, refs_dir)
