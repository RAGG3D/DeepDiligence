#!/usr/bin/env python3
"""Fill Historical Events prices, DoD moves, EVT/category cells.

The template stores four year blocks. Each block still has an old Bloomberg
pull/helper area plus a clean event-study table.  This script rewrites the clean
table as:

    Date | Share Price | DoD Chg | EVT | Category

It also hides the stale Bloomberg/helper columns, fills prices from public daily
closes, and supplements approved company events with >=8% daily share-price
moves so large unexplained moves are not silently missed.
"""
import argparse
import csv
import json
import re
import shutil
import zipfile
from io import BytesIO
from datetime import date, datetime
from datetime import timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlparse
import xml.etree.ElementTree as ET

from openpyxl.utils.datetime import to_excel
from openpyxl.utils import column_index_from_string, get_column_letter

try:
    import requests
    from bs4 import BeautifulSoup
except Exception:  # pragma: no cover - optional runtime dependency
    requests = None
    BeautifulSoup = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

try:
    import yfinance as yf
except Exception:  # pragma: no cover - dependency availability is environment-specific
    yf = None

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", _NS_MAIN)
ET.register_namespace("r", _NS_R)

BLOCKS = [
    # header, legacy_date, event_date_col, share_price_col, dod_col, evt_col, category_col
    ("B7", "B9", "G", "H", "I", "J", "K"),
    ("M7", "M9", "R", "S", "T", "U", "V"),
    ("X7", "X9", "AC", "AD", "AE", "AF", "AG"),
    ("AI7", "AI9", "AN", "AO", "AP", "AQ", "AR"),
]

LEGACY_COLUMN_RANGES = [(3, 6), (14, 17), (25, 28), (36, 39)]

VALID_CATEGORIES = {
    "Clinical Data", "Partnership", "Regulatory", "Financing", "Corporate", "Other",
}

CONFERENCE_RE = re.compile(
    r"\b(ASCO(?:\s+GU|\s+GI)?|AACR|ESMO(?:\s+IO)?|SITC|ASH|EHA|ISSVA|"
    r"SABCS|SNO|WCLC|AAN|CTAD|ACR|EULAR|AASLD|DDW|ATS|ERS|ASN|AHA|"
    r"ACC|ESC|AAO|ARVO|CROI|IDWEEK|WORLDSYMPOSIUM|ACMG|ASHG)\b",
    re.I,
)

CLINICAL_DATA_RE = re.compile(
    r"\b(phase\s*[1-4]|clinical\s+(?:data|results)|interim|updated|primary|"
    r"readout|patient[s]?|efficacy|safety|ORR|CR\b|PFS|OS\b|survival)\b",
    re.I,
)


def _xml_escape(text: str) -> str:
    return (
        str(text).replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _col_of(addr: str) -> int:
    return column_index_from_string("".join(c for c in addr if c.isalpha()))


def _row_of(addr: str) -> int:
    return int("".join(c for c in addr if c.isdigit()))


def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_path: Dict[str, str] = {}
    for rel in rels_xml:
        if "worksheet" in rel.get("Type", ""):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rid_to_path[rid] = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")

    for sheet in wb_xml.findall(f".//{{{_NS_MAIN}}}sheet"):
        name = sheet.get("name", "")
        rid = sheet.get(f"{{{_NS_R}}}id", "")
        if name == sheet_name and rid in rid_to_path:
            return rid_to_path[rid]
    return None


def _cell_bounds(xml: str, addr: str) -> Optional[Tuple[int, int, int, int]]:
    search = f'r="{addr}"'
    pos = xml.find(search)
    if pos == -1:
        return None
    lt = xml.rfind("<", 0, pos)
    if lt == -1 or xml[lt + 1] != "c":
        return None
    tag_end = xml.index(">", lt) + 1
    if xml[tag_end - 2:tag_end] == "/>":
        return lt, tag_end, tag_end - 2, tag_end
    c_end = xml.index("</c>", tag_end) + 4
    return lt, tag_end, c_end, c_end


def _insert_row(xml: str, row_num: int, cell_xml: str) -> str:
    """Create a ``<row r="N">`` holding ``cell_xml``, inserted into ``<sheetData>``
    in ascending row order (before the first existing row with a larger index,
    or before ``</sheetData>`` when none is larger).

    The template's four year-blocks only ship 365 data rows (9-373); Dec 31 of a
    leap year lands on row 374, which does not exist yet, so it must be created."""
    new_row = f'<row r="{row_num}">{cell_xml}</row>'
    for m in re.finditer(r'<row\b[^>]*\br="(\d+)"', xml):
        if int(m.group(1)) > row_num:
            return xml[:m.start()] + new_row + xml[m.start():]
    idx = xml.find("</sheetData>")
    if idx == -1:
        raise ValueError("sheetData not found")
    return xml[:idx] + new_row + xml[idx:]


def _insert_cell(xml: str, row_num: int, addr: str, cell_xml: str) -> str:
    row_search = f'r="{row_num}"'
    pos = 0
    lt = -1
    while True:
        rp = xml.find(row_search, pos)
        if rp == -1:
            # Row absent from the grid (e.g. Dec 31 of a leap year): create it.
            return _insert_row(xml, row_num, cell_xml)
        cand = xml.rfind("<", 0, rp)
        if cand != -1 and xml[cand + 1:cand + 4] == "row":
            lt = cand
            break
        pos = rp + 1
    row_tag_end = xml.index(">", lt) + 1
    if xml[row_tag_end - 2:row_tag_end] == "/>":
        return xml[:row_tag_end - 2] + f">{cell_xml}</row>" + xml[row_tag_end:]

    row_end = xml.index("</row>", row_tag_end)
    row_body = xml[row_tag_end:row_end]
    col_idx = _col_of(addr)
    insert_at = len(row_body)
    for m in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"', row_body):
        if _col_of(m.group(1)) > col_idx:
            insert_at = m.start()
            break
    return xml[:row_tag_end] + row_body[:insert_at] + cell_xml + row_body[insert_at:] + xml[row_end:]


def _patch_text_cell(xml: str, addr: str, text: str) -> str:
    escaped = _xml_escape(text)
    row_num = _row_of(addr)
    bounds = _cell_bounds(xml, addr)
    if bounds is None:
        try:
            return _insert_cell(
                xml, row_num, addr,
                f'<c r="{addr}" t="inlineStr"><is><t>{escaped}</t></is></c>',
            )
        except ValueError:
            return xml
    lt, tag_end, c_end, _ = bounds
    self_closing = c_end < tag_end
    open_tag = xml[lt:tag_end - 2] if self_closing else xml[lt:tag_end - 1]
    open_tag = re.sub(r'\s+t="[^"]*"', "", open_tag)
    tail_start = tag_end if self_closing else c_end
    return (
        xml[:lt]
        + open_tag
        + f' t="inlineStr"><is><t>{escaped}</t></is></c>'
        + xml[tail_start:]
    )


def _patch_number_cell(xml: str, addr: str, value: float) -> str:
    row_num = _row_of(addr)
    val = str(int(value)) if value == int(value) else str(value)
    bounds = _cell_bounds(xml, addr)
    if bounds is None:
        try:
            return _insert_cell(xml, row_num, addr, f'<c r="{addr}"><v>{val}</v></c>')
        except ValueError:
            return xml
    lt, tag_end, c_end, _ = bounds
    if c_end < tag_end:
        open_tag = xml[lt:tag_end - 2]
        tail_start = tag_end
    else:
        open_tag = xml[lt:tag_end - 1]
        tail_start = c_end
    open_tag = re.sub(r'\s+t="[^"]*"', "", open_tag)
    return xml[:lt] + open_tag + f"><v>{val}</v></c>" + xml[tail_start:]


def _patch_calc_mode(parts: Dict[str, bytes], xlsx_path: Path) -> None:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        ct = zf.read("[Content_Types].xml").decode("utf-8")
        wr = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', "", ct)
    wr = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", wr)
    parts["xl/workbook.xml"] = wb_xml.encode("utf-8")
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")


def _sheet_rels_path(sheet_zip: str) -> str:
    parent, name = sheet_zip.rsplit("/", 1)
    return f"{parent}/_rels/{name}.rels"


def _patch_hyperlinks(
    xml: str,
    rels_xml: Optional[str],
    links: Dict[str, str],
) -> Tuple[str, str]:
    """Replace Historical Events cell hyperlinks while preserving other links."""
    if rels_xml:
        # Old EVT links are rebuilt below. Remove only hyperlink relationships;
        # drawings/comments and all other worksheet relationships survive.
        rels_xml = re.sub(
            r'<Relationship\b[^>]*Type="[^"]*/hyperlink"[^>]*/>', "", rels_xml
        )
    else:
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>'
        )

    used = [int(x) for x in re.findall(r'\bId="rId(\d+)"', rels_xml)]
    next_id = max(used, default=0) + 1
    hyperlink_nodes: List[str] = []
    relationship_nodes: List[str] = []
    for addr, url in sorted(links.items(), key=lambda item: (_row_of(item[0]), _col_of(item[0]))):
        if not url.startswith(("http://", "https://")):
            continue
        rid = f"rId{next_id}"
        next_id += 1
        hyperlink_nodes.append(f'<hyperlink ref="{addr}" r:id="{rid}"/>')
        relationship_nodes.append(
            '<Relationship '
            f'Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{_xml_escape(url)}" TargetMode="External"/>'
        )

    xml = re.sub(r'<hyperlinks>.*?</hyperlinks>', "", xml, flags=re.S)
    if hyperlink_nodes:
        block = "<hyperlinks>" + "".join(hyperlink_nodes) + "</hyperlinks>"
        for marker in ("<printOptions", "<pageMargins", "<pageSetup", "</worksheet>"):
            pos = xml.find(marker)
            if pos != -1:
                xml = xml[:pos] + block + xml[pos:]
                break
    if relationship_nodes:
        rels_xml = rels_xml.replace(
            "</Relationships>", "".join(relationship_nodes) + "</Relationships>"
        )
    return xml, rels_xml


def _parse_date(value: Any) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            pass
    return None


def _load_events(path: Path) -> List[Dict[str, str]]:
    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))
        return data.get("events", data) if isinstance(data, dict) else data
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))

    events: List[Dict[str, str]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3 or cells[0].lower() in {"date", "---"}:
            continue
        if not re.match(r"\d{4}-\d{2}-\d{2}", cells[0]):
            continue
        events.append({"date": cells[0], "evt": cells[1], "category": cells[2]})
    return events


def _normalize_event(event: Dict[str, Any]) -> Optional[Dict[str, str]]:
    d = _parse_date(event.get("date") or event.get("Date"))
    if not d:
        return None
    evt = (event.get("evt") or event.get("EVT") or "").strip()
    if not evt:
        return None
    category = (event.get("category") or event.get("Category") or "").strip()
    if category not in VALID_CATEGORIES:
        category = "Other"
    out = {"date": d.isoformat(), "evt": evt, "category": category}
    for key in (
        "url", "source_url", "abstract_url", "venue", "source_kind",
        "phase", "data_type", "patient_size", "orr", "cr", "survival", "safety",
    ):
        value = event.get(key)
        if value not in (None, ""):
            out[key] = str(value).strip()
    if "url" not in out and "source_url" in out:
        out["url"] = out["source_url"]
    return out


def _venue(text: str) -> str:
    match = CONFERENCE_RE.search(text or "")
    if match:
        return re.sub(r"\s+", " ", match.group(1)).upper()
    low = (text or "").lower()
    if "conference call" in low or "earnings call" in low:
        return "CALL"
    if "investor day" in low:
        return "INVESTOR DAY"
    return "COMPANY PR"


def _metric_value(text: str, patterns: Iterable[str]) -> Optional[str]:
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return re.sub(r"\s+", " ", match.group(1)).strip(" ,;.")
    return None


def _structured_clinical_summary(title: str, body: str, abstract_body: str = "") -> str:
    """Create a dense, source-bounded clinical-data line.

    This intentionally extracts only explicit numbers. Missing metrics are
    omitted instead of guessed. Conference venue is always the first token.
    """
    combined = re.sub(r"\s+", " ", " ".join([title, body, abstract_body])).strip()
    venue = _venue(combined)
    phase = _metric_value(combined, [r"\b((?:Phase|PHASE)\s*[1-4](?:[abAB]|b/3|2/3)?)\b"])
    data_type = _metric_value(combined, [
        r"\b((?:initial|interim|updated|primary|final|preliminary)\s+(?:clinical\s+)?(?:data|results))\b",
        r"\b((?:initial|interim|updated|primary|final|preliminary))\b",
    ])
    n_value = _metric_value(combined, [
        r"\b((?:N|n)\s*=\s*\d+)\b",
        r"\b((?:\d+)\s+(?:evaluable|treated|enrolled|dosed)\s+patients?)\b",
        r"\b((?:patient size|sample size)\s*(?:of|=|:)\s*\d+)\b",
    ])
    orr = _metric_value(combined, [
        r"\b((?:ORR|overall response rate)\s*(?:was|of|=|:)?\s*\d+(?:\.\d+)?%)",
    ])
    cr = _metric_value(combined, [
        r"\b((?:CR|complete response(?: rate)?)\s*(?:was|of|=|:)?\s*\d+(?:\.\d+)?%)",
        r"\b((?:\d+\s*/\s*\d+)\s+(?:patients?\s+)?(?:achieved|had)\s+(?:a\s+)?complete response)",
    ])
    survival = _metric_value(combined, [
        r"\b((?:mPFS|median PFS|PFS|mOS|median OS|OS)\s*(?:was|of|=|:)?\s*\d+(?:\.\d+)?\s*(?:months?|mos?|%))",
        r"\b((?:\d+(?:\.\d+)?[- ]month\s+(?:PFS|OS|survival))\s*(?:was|of|=|:)?\s*\d+(?:\.\d+)?%)",
    ])
    safety = _metric_value(combined, [
        r"\b((?:no\s+)?(?:Grade|grade)\s*[3-5](?:\+)?[^.;]{0,80}(?:AE|AEs|TEAE|TEAEs|TRAE|TRAEs|adverse events?)[^.;]{0,50})",
        r"\b((?:SAE|SAEs|serious adverse events?|dose[- ]limiting toxicit(?:y|ies))[^.;]{0,80}\d+(?:\.\d+)?%)",
        r"\b((?:well tolerated|generally well tolerated|no new safety signals))\b",
    ])
    fields = [x for x in [phase, data_type, n_value, orr, cr, survival, safety] if x]
    if not fields:
        fields = [_short_event_title(title, max_words=18)]
    return f"[{venue}] " + "; ".join(fields)


def _abstract_link(soup: "Any", base_url: str) -> Optional[str]:
    candidates: List[Tuple[int, str]] = []
    base_host = urlparse(base_url).netloc.lower()
    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a.get("href", "")).split("#", 1)[0]
        label = (a.get_text(" ", strip=True) + " " + href).lower()
        if not href.startswith(("http://", "https://")):
            continue
        target_host = urlparse(href).netloc.lower()
        if target_host == base_host and "/news-releases/news-release-details/" in href.lower():
            continue
        score = 0
        if "abstract" in (a.get_text(" ", strip=True) or "").lower():
            score += 5
        if any(x in label for x in ("poster", "presentation", "scientific-program")):
            score += 3
        if any(x in urlparse(href).netloc.lower() for x in (
            "asco.org", "aacr.org", "ashpublications.org", "esmo.org",
            "sitcancer.org", "issva.org", "ehaweb.org",
        )):
            score += 4
        if href.lower().split("?", 1)[0].endswith(".pdf"):
            score += 1
        same_site_document = (
            href.lower().split("?", 1)[0].endswith(".pdf")
            or "/static-files/" in href.lower()
        )
        if target_host == base_host and not same_site_document:
            score = 0
        if score:
            candidates.append((score, href))
    return max(candidates, default=(0, ""))[1] or None


def _categorize_news(title: str) -> str:
    low = title.lower()
    if any(x in low for x in [
        "phase", "clinical", "patient", "data", "trial", "readout",
        "poster", "clinical results", "response rate", "complete response",
    ]):
        return "Clinical Data"
    if any(x in low for x in ["partner", "collaboration", "agreement", "license", "alliance"]):
        return "Partnership"
    if any(x in low for x in ["fda", "ind ", "regulatory", "approval", "approved"]):
        return "Regulatory"
    if any(x in low for x in ["financing", "offering", "placement", "raise", "funding"]):
        return "Financing"
    if any(x in low for x in [
        "financial results", "highlights", "annual general meeting", "board",
        "appoint", "cash runway", "operational efficiencies", "management",
        "reports q", "reports full year", "invitation",
    ]):
        return "Corporate"
    return "Other"


def _short_event_title(title: str, company_name: Optional[str] = None, max_words: int = 15) -> str:
    text = title.strip()
    if company_name:
        # Strip a leading occurrence of the issuer's own name so EVT titles are
        # not prefixed with it. Try the full legal name first, then the name with
        # corporate suffixes (AG/Inc./Ltd./...) trimmed, since titles use either.
        full = company_name.strip()
        short = re.sub(
            r"[\s,]+(?:AG|N\.?V\.?|S\.?A\.?|Inc\.?|Incorporated|Corp\.?|Corporation|"
            r"Ltd\.?|Limited|LLC|PLC|Co\.?|Holdings?|Group)\.?$",
            "",
            full,
            flags=re.I,
        ).strip()
        for name in (full, short):
            if not name:
                continue
            stripped = re.sub(rf"^{re.escape(name)}[\s,]+", "", text, flags=re.I).strip()
            if stripped != text:
                text = stripped
                break
    text = re.sub(r"\s+", " ", text)
    words = text.split()
    if len(words) > max_words:
        text = " ".join(words[:max_words])
    return text.strip(" -") or "company news"


def _parse_news_date(text: str) -> Optional[date]:
    text = re.sub(r"\s+", " ", text).strip()
    for fmt in ("%B %d, %Y", "%b %d, %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:30], fmt).date()
        except ValueError:
            pass
    m = re.search(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}",
        text,
    )
    if m:
        try:
            return datetime.strptime(m.group(0), "%B %d, %Y").date()
        except ValueError:
            return None
    m = re.search(
        r"\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|"
        r"Dec(?:ember)?)\.?\s+(\d{1,2}),\s+(\d{4})\b",
        text,
        re.I,
    )
    if m:
        token = m.group(1).lower()[:3].title()
        try:
            return datetime.strptime(f"{token} {m.group(2)} {m.group(3)}", "%b %d %Y").date()
        except ValueError:
            return None
    return None


def _is_candidate_news_link(
    abs_href: str,
    base_host: str,
    base_path: str,
    link_re: "Optional[re.Pattern[str]]",
) -> bool:
    """Decide whether an <a> href is a plausible news-detail link.

    An explicit ``link_re`` (from --news-link-pattern) is authoritative. Otherwise
    the Q4/GlobeNewswire detail path is accepted as one explicit case, and ANY
    same-site article link is kept — either one that descends below the news-index
    path, or one whose path carries a news/press/release-ish segment plus a
    slug-like tail. The detail page's in-window date parse (below) is the real
    validator, so a generous match here only costs an extra fetch, never a wrong
    event."""
    low = abs_href.lower()
    if low.startswith(("mailto:", "javascript:", "tel:")):
        return False
    if link_re is not None:
        return bool(link_re.search(abs_href))
    # Keep the Q4/GlobeNewswire detail path working as one explicit case.
    if "/news-releases/news-release-details/" in low:
        return True
    if low.rsplit("?", 1)[0].endswith(
        (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".zip", ".xlsx", ".doc", ".docx")
    ):
        return False
    host = urlparse(abs_href).netloc.lower()
    if host and base_host and host != base_host:
        return False
    path = urlparse(abs_href).path.rstrip("/")
    if not path:
        return False
    # Same-site link that descends below the news-index path (original behavior).
    if base_path and path.startswith(base_path) and len(path) > len(base_path):
        return True
    # Generalized: any same-site link that looks like an individual news/press
    # article — an article-ish path segment plus a slug-like tail below it, so
    # nav/index links (a bare "/news") are skipped.
    low_path = path.lower()
    segments = [s for s in low_path.split("/") if s]
    if len(segments) >= 2 and any(
        hint in low_path for hint in ("news", "press", "release", "announce", "media")
    ):
        return True
    return False


def _fetch_official_news_events(
    news_url: Optional[str],
    years: Optional[List[int]],
    company_name: Optional[str] = None,
    date_selector: Optional[str] = None,
    link_pattern: Optional[str] = None,
    page_template: Optional[str] = None,
) -> List[Dict[str, str]]:
    if not news_url:
        return []
    if requests is None or BeautifulSoup is None:
        print("  WARNING: official-news scrape skipped for "
              f"{news_url} — requests/bs4 unavailable in this environment.")
        return []
    base_url = news_url
    base_host = urlparse(base_url).netloc.lower()
    base_path = urlparse(base_url).path.rstrip("/")
    # Pagination shape defaults to the Q4/GlobeNewswire "?page=N" template.
    page_template = page_template or "?page={page}"
    link_re = re.compile(link_pattern) if link_pattern else None

    min_year = min(years) if years else None
    max_year = max(years) if years else None
    session = requests.Session()
    headers = {"User-Agent": "DeepDiligence event updater/1.0"}
    detail_links: list[tuple[str, str]] = []
    seen_links: set[str] = set()

    consecutive_empty_pages = 0
    for page in range(0, 30):
        url = base_url if page == 0 else f"{base_url}{page_template.format(page=page)}"
        try:
            resp = session.get(url, headers=headers, timeout=20)
        except Exception:
            break
        if resp.status_code >= 400:
            break
        soup = BeautifulSoup(resp.text, "html.parser")
        page_links = 0
        for a in soup.find_all("a", href=True):
            title = a.get_text(" ", strip=True)
            href = a["href"]
            if not title or title.lower() == "pdf version":
                continue
            href = urljoin(base_url, href)
            href = href.split("#", 1)[0].rstrip("/")
            if not _is_candidate_news_link(href, base_host, base_path, link_re):
                continue
            if href in seen_links:
                continue
            seen_links.add(href)
            detail_links.append((title, href))
            page_links += 1
        if page_links == 0:
            # Some issuer-search endpoints expose page 0 twice (base URL and
            # page=1) before older results begin at page=2.  Do not truncate the
            # archive on that single duplicate page; two consecutive pages with
            # no new articles still terminate malformed/infinite pagination.
            consecutive_empty_pages += 1
            if consecutive_empty_pages >= 2:
                break
            continue
        consecutive_empty_pages = 0
        # Four rolling years can easily exceed 200 releases for an active
        # issuer. Keep crawling until pagination ends; this high guard only
        # protects against a malformed site that generates infinite URLs.
        if len(detail_links) >= 1000:
            break

    if not detail_links:
        print(f"  WARNING: official-news scrape found 0 article links at {base_url}")
        return []

    events: list[Dict[str, str]] = []
    for link_title, href in detail_links:
        d: Optional[date] = None
        dated_path = re.search(r"/news-release/(\d{4})/(\d{2})/(\d{2})/", href)
        if dated_path:
            try:
                d = date(*(int(dated_path.group(i)) for i in range(1, 4)))
            except ValueError:
                d = None
        title = link_title
        body = ""
        abstract_url: Optional[str] = None
        abstract_body = ""
        try:
            resp = session.get(href, headers=headers, timeout=20)
            if resp.status_code < 400:
                soup = BeautifulSoup(resp.text, "html.parser")
                og_title = soup.find("meta", attrs={"property": "og:title"})
                if og_title and og_title.get("content"):
                    title = str(og_title["content"]).split(" | ", 1)[0].strip()
                heading = soup.find("h1")
                heading_text = heading.get_text(" ", strip=True) if heading else ""
                if (not og_title) and heading_text and heading_text.lower() not in {"press release", "news release", "news releases"}:
                    title = heading_text
                main = soup.find("article") or soup.find("main") or soup.body
                if main:
                    body = main.get_text(" ", strip=True)
                date_el = soup.select_one(date_selector) if date_selector else None
                if d is None and date_el:
                    d = _parse_news_date(date_el.get_text(" ", strip=True))
                if d is None:
                    # Prefer the issuer-visible release date. UTC metadata can
                    # shift an after-hours U.S. release into the following day,
                    # which breaks event/price alignment and date completeness.
                    published = soup.select_one(
                        "time[itemprop='datePublished'], .article-published time, "
                        "[itemprop='datePublished']"
                    )
                    if published:
                        d = _parse_news_date(published.get_text(" ", strip=True))
                        if d is None and published.get("datetime"):
                            d = _parse_date(str(published.get("datetime"))[:10])
                if d is None:
                    for meta_name in ["article:published_time", "og:updated_time", "date"]:
                        meta = soup.find("meta", attrs={"property": meta_name}) or soup.find("meta", attrs={"name": meta_name})
                        if meta and meta.get("content"):
                            d = _parse_date(meta["content"][:10])
                            if d:
                                break
                if d is None:
                    for node in soup.find_all(["p", "div", "span"]):
                        node_text = node.get_text(" ", strip=True)
                        if "GLOBE NEWSWIRE" in node_text.upper() or "BUSINESS WIRE" in node_text.upper():
                            d = _parse_news_date(node_text[:500])
                            if d:
                                break
                if d is None:
                    for node in soup.find_all(["time", "p", "span"], limit=80):
                        d = _parse_news_date(node.get_text(" ", strip=True))
                        if d:
                            break
                abstract_url = _abstract_link(soup, href)
                if abstract_url:
                    try:
                        ar = session.get(abstract_url, headers=headers, timeout=20)
                        ctype = ar.headers.get("content-type", "").lower()
                        if ar.status_code < 400 and "html" in ctype:
                            abstract_body = BeautifulSoup(ar.text, "html.parser").get_text(" ", strip=True)
                        elif ar.status_code < 400 and PdfReader is not None and (
                            "pdf" in ctype or ar.content[:4] == b"%PDF"
                        ):
                            reader = PdfReader(BytesIO(ar.content))
                            abstract_body = " ".join((page.extract_text() or "") for page in reader.pages)
                    except Exception:
                        abstract_body = ""
        except Exception:
            d = None
        if d is None:
            continue
        if min_year and d.year < min_year:
            continue
        if max_year and d.year > max_year:
            continue
        # Body boilerplate nearly always mentions the issuer's Phase 2 trials,
        # so clinical classification must be driven by the release headline.
        # Detailed approved-research records can still override this later.
        is_clinical = _categorize_news(title) == "Clinical Data"
        evt = (
            _structured_clinical_summary(title, body, abstract_body)
            if is_clinical
            else _short_event_title(title, company_name)
        )
        event: Dict[str, str] = {
            "date": d.isoformat(),
            "title": title,
            "evt": evt,
            "category": "Clinical Data" if is_clinical else _categorize_news(title),
            "url": href,
            "source_url": href,
            "source_kind": "company_press_release",
            "venue": _venue(" ".join([title, body])),
        }
        if abstract_url:
            event["abstract_url"] = abstract_url
        events.append(event)

    # Deduplicate by date+EVT while preserving chronological order.
    out: list[Dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for event in sorted(events, key=lambda e: (e["date"], e["evt"])):
        key = (event["date"], event.get("url", event["evt"]).lower())
        if key in seen:
            continue
        seen.add(key)
        out.append(event)
    if out:
        print(f"  official news events fetched: {len(out)} from {base_url}")
    return out


def _download_closes(ticker: str, start: date, end: date) -> Dict[date, float]:
    if yf is None:
        print("  yfinance unavailable; Historical Events prices were not refreshed")
        return {}
    if end < start:
        return {}
    try:
        df = yf.download(
            ticker,
            start=start.isoformat(),
            end=(end + timedelta(days=1)).isoformat(),
            progress=False,
            auto_adjust=False,
            threads=False,
        )
    except Exception as exc:
        print(f"  yfinance download failed for {ticker}: {exc}")
        return {}
    if df is None or df.empty or "Close" not in df:
        print(f"  yfinance returned no Close history for {ticker}")
        return {}
    close = df["Close"]
    if hasattr(close, "columns"):
        close = close.iloc[:, 0]
    out: Dict[date, float] = {}
    for idx, value in close.dropna().items():
        if hasattr(idx, "date"):
            d = idx.date()
        else:
            d = _parse_date(idx)
        if d and start <= d <= end:
            out[d] = float(value)
    return out


def _calendar_price_series(
    closes: Dict[date, float],
    start: date,
    end: date,
) -> Tuple[Dict[date, float], Dict[date, float], Dict[date, float]]:
    """Return filled calendar prices, calendar DoD, and trading-day DoD."""
    filled: Dict[date, float] = {}
    calendar_dod: Dict[date, float] = {}
    trading_dod: Dict[date, float] = {}
    if not closes:
        return filled, calendar_dod, trading_dod

    trading_days = sorted(closes)
    last_available = trading_days[-1]
    previous_close: Optional[float] = None
    for d in trading_days:
        close = closes[d]
        trading_dod[d] = 0.0 if previous_close in (None, 0) else close / previous_close - 1.0
        previous_close = close

    carry: Optional[float] = None
    d = start
    while d <= end:
        if d > last_available:
            d += timedelta(days=1)
            continue
        if d in closes:
            carry = closes[d]
            filled[d] = carry
            calendar_dod[d] = trading_dod[d]
        elif carry is not None:
            filled[d] = carry
            calendar_dod[d] = 0.0
        d += timedelta(days=1)
    return filled, calendar_dod, trading_dod


def _nearest_event(
    events: Iterable[Dict[str, str]],
    target: date,
    max_days: int = 2,
) -> Optional[Dict[str, str]]:
    best: Optional[Tuple[int, Dict[str, str]]] = None
    for event in events:
        d = _parse_date(event.get("date"))
        if not d:
            continue
        delta = abs((target - d).days)
        if delta > max_days:
            continue
        if best is None or delta < best[0]:
            best = (delta, event)
    return best[1] if best else None


def _market_move_label(target: date, market_moves: Dict[str, Dict[date, float]]) -> str:
    labels = []
    # A ~2.5% SPY day is already a genuine broad-market move, so use a lower
    # threshold there; keep the higher 4% bar for the more volatile biotech ETF.
    for symbol, name, threshold in (("SPY", "market", 0.025), ("XBI", "biotech", 0.04)):
        move = market_moves.get(symbol, {}).get(target)
        if move is not None and abs(move) >= threshold:
            labels.append(name)
    if not labels:
        return ""
    return "/".join(labels)


def _shorten_evt(text: str, max_words: int = 10) -> str:
    words = re.findall(r"[A-Za-z0-9$%+\-.]+", text)
    if not words:
        return "company update"
    return " ".join(words[:max_words])


def _load_move_research(ticker: Optional[str]) -> Dict[str, Dict[str, str]]:
    """Load researched causes for large-move dates from a prior research pass.

    A two-pass flow writes unexplained >=threshold dates to
    ``{ticker}_moves_needing_research.json``; a research follow-up resolves them
    into ``{ticker}_moves_researched.json`` ({"moves": [{date, evt, category}]}),
    which this reads so a real cause is preferred over the generic filler."""
    if not ticker:
        return {}
    path = Path("artifacts") / ticker / f"{ticker}_moves_researched.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    items = data.get("moves", data) if isinstance(data, dict) else data
    out: Dict[str, Dict[str, str]] = {}
    for item in items or []:
        if not isinstance(item, dict):
            continue
        d = _parse_date(item.get("date") or item.get("Date"))
        evt = (item.get("evt") or item.get("EVT") or "").strip()
        if d and evt:
            out[d.isoformat()] = {
                "evt": evt,
                "category": (item.get("category") or item.get("Category") or "Other").strip(),
            }
    return out


def _write_moves_needing_research(ticker: Optional[str], moves: List[Dict[str, Any]]) -> None:
    if not ticker:
        return
    out_dir = Path("artifacts") / ticker
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{ticker}_moves_needing_research.json"
    out_path.write_text(json.dumps({"moves": moves}, indent=2), encoding="utf-8")
    if moves:
        print(f"  {len(moves)} large moves need cause research → {out_path}")


def _add_large_move_events(
    events: List[Dict[str, str]],
    trading_dod: Dict[date, float],
    threshold: float,
    market_moves: Dict[str, Dict[date, float]],
    ticker: Optional[str] = None,
) -> List[Dict[str, str]]:
    existing_dates = {_parse_date(e.get("date")) for e in events}
    enriched = list(events)
    researched = _load_move_research(ticker)
    needing_research: List[Dict[str, Any]] = []
    for d, move in sorted(trading_dod.items()):
        if abs(move) < threshold:
            continue
        if d in existing_dates:
            continue
        near = _nearest_event(events, d)
        if near:
            evt = f"Share move after {_shorten_evt(near['evt'], 7)}"
            # This row explains a different trading date's price move; it is not
            # itself a clinical-data disclosure and therefore must not inherit
            # the source event's Clinical Data category/hyperlink requirements.
            category = "Other"
        else:
            market_label = _market_move_label(d, market_moves)
            if market_label:
                evt = f"Share move with {market_label} weakness" if move < 0 else f"Share move with {market_label} rally"
                category = "Other"
            else:
                cause = researched.get(d.isoformat())
                if cause and cause.get("evt"):
                    # A prior research pass explained this move; use its cause.
                    evt = cause["evt"]
                    category = cause.get("category") or "Other"
                else:
                    # No cause available yet: flag the date for a research
                    # follow-up and fall back to the generic filler for now.
                    needing_research.append(
                        {"date": d.isoformat(), "dod_pct": round(move * 100, 2)}
                    )
                    evt = "Large share move; no company news found"
                    category = "Other"
        enriched.append({"date": d.isoformat(), "evt": evt[:90], "category": category})
        existing_dates.add(d)
    _write_moves_needing_research(ticker, needing_research)
    return sorted(enriched, key=lambda e: e["date"])


def _hide_legacy_columns(xml: str) -> str:
    def repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        min_m = re.search(r'\bmin="(\d+)"', tag)
        max_m = re.search(r'\bmax="(\d+)"', tag)
        if not min_m or not max_m:
            return tag
        col_min = int(min_m.group(1))
        col_max = int(max_m.group(1))
        if not any(col_min <= hi and col_max >= lo for lo, hi in LEGACY_COLUMN_RANGES):
            return tag
        tag = re.sub(r'\bhidden="[^"]*"', 'hidden="true"', tag)
        if 'hidden="' not in tag:
            tag = tag.replace("<col ", '<col hidden="true" ', 1)
        tag = re.sub(r'\bwidth="[^"]*"', 'width="0"', tag)
        if 'width="' not in tag:
            tag = tag.replace("/>", ' width="0"/>')
        tag = re.sub(r'\bcustomWidth="[^"]*"', 'customWidth="true"', tag)
        if 'customWidth="' not in tag:
            tag = tag.replace("<col ", '<col customWidth="true" ', 1)
        return tag

    return re.sub(r"<col\b[^>]*/>", repl, xml)


def _write_enriched_events(ticker: Optional[str], events: List[Dict[str, str]]) -> None:
    if not ticker:
        return
    out_dir = Path("artifacts") / ticker
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{ticker}_historical_events_enriched.json"
    official_urls = sorted({
        str(e.get("url") or e.get("source_url"))
        for e in events
        if (
            e.get("source_kind") == "company_press_release"
            or "/news-releases/news-release-details/" in str(
                e.get("url") or e.get("source_url") or ""
            )
        )
    })
    payload = {
        "official_release_count": len(official_urls),
        "official_release_urls": official_urls,
        "events": events,
    }
    out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(f"  enriched events → {out_path}")


def _events_by_date(events: Iterable[Dict[str, str]]) -> List[Dict[str, str]]:
    """Combine all same-day releases into the single worksheet EVT cell.

    The artifact retains every release as its own record. Only the display layer
    is combined because the template provides one EVT cell per calendar day.
    """
    grouped: Dict[str, List[Dict[str, str]]] = {}
    for event in events:
        grouped.setdefault(event["date"], []).append(event)
    priority = {
        "Clinical Data": 0, "Regulatory": 1, "Financing": 2,
        "Partnership": 3, "Corporate": 4, "Other": 5,
    }
    out: List[Dict[str, str]] = []
    for d, items in sorted(grouped.items()):
        items = sorted(items, key=lambda e: (priority.get(e.get("category", "Other"), 9), e["evt"]))
        unique_text: List[str] = []
        seen_text: set[str] = set()
        for item in items:
            text = item["evt"].strip()
            key = text.lower()
            if key not in seen_text:
                unique_text.append(text)
                seen_text.add(key)
        primary = items[0]
        link = next(
            (e.get("abstract_url") for e in items if e.get("abstract_url")),
            None,
        ) or next((e.get("url") or e.get("source_url") for e in items if e.get("url") or e.get("source_url")), None)
        combined = {
            "date": d,
            "evt": " | ".join(unique_text),
            "category": primary.get("category", "Other"),
        }
        if link:
            combined["url"] = link
        out.append(combined)
    return out


def _year_span(events: Iterable[Dict[str, str]], explicit: Optional[List[int]]) -> List[int]:
    if explicit:
        return explicit[:4]
    years = sorted({
        _parse_date(e.get("date") or e.get("Date")).year
        for e in events
        if _parse_date(e.get("date") or e.get("Date"))
    })
    if not years:
        last = date.today().year
        return [last - 3, last - 2, last - 1, last]
    while len(years) < 4:
        years.insert(0, years[0] - 1)
    return years[-4:]


def fill_historical_events(
    xlsx_path: Path,
    events_path: Optional[Path] = None,
    ticker: Optional[str] = None,
    years: Optional[List[int]] = None,
    include_price_moves: bool = True,
    move_threshold: float = 0.08,
    include_official_news: bool = True,
    news_url: Optional[str] = None,
    company_name: Optional[str] = None,
    news_date_selector: Optional[str] = None,
    news_link_pattern: Optional[str] = None,
    news_page_template: Optional[str] = None,
    require_official_news: bool = False,
) -> int:
    sheet_zip = _get_sheet_zip_path(xlsx_path, "Historical Events")
    if not sheet_zip:
        raise SystemExit("Historical Events sheet not found")
    raw_events = _load_events(events_path) if events_path and events_path.exists() else []
    events = [event for event in (_normalize_event(e) for e in raw_events) if event]
    block_years = _year_span(events, years)
    if include_official_news:
        official = _fetch_official_news_events(
            news_url, block_years, company_name, news_date_selector,
            news_link_pattern, news_page_template,
        )
        if require_official_news and not official:
            raise RuntimeError(
                "Official press-release audit returned zero releases; refusing to deliver "
                "a Historical Events sheet built only from secondary/approved snippets"
            )
        if official:
            # If an approved research record and exactly one official PR share
            # date+category, retain the analyst's richer wording and attach the
            # official source lineage. Multiple same-day PRs remain separate.
            approved_buckets: Dict[Tuple[str, str], List[Dict[str, str]]] = {}
            official_buckets: Dict[Tuple[str, str], List[Dict[str, str]]] = {}
            for event in events:
                approved_buckets.setdefault((event["date"], event["category"]), []).append(event)
            for event in official:
                official_buckets.setdefault((event["date"], event["category"]), []).append(event)
            consumed: set[int] = set()
            for key, approved_items in approved_buckets.items():
                official_items = official_buckets.get(key, [])
                if len(approved_items) == len(official_items) == 1:
                    approved = approved_items[0]
                    source = official_items[0]
                    # A single approved research record and a single issuer PR
                    # on the same date/category form an unambiguous pair.  The
                    # approved wording is often deliberately metric-only and may
                    # omit every title token (e.g. "[CALL] PFS/OS..."); requiring
                    # lexical overlap would preserve the poorer auto-summary.
                    for field in ("url", "source_url", "abstract_url", "venue", "source_kind"):
                        if source.get(field) and not approved.get(field):
                            approved[field] = source[field]
                    consumed.add(id(source))
            merged: dict[tuple[str, str], Dict[str, str]] = {
                (e["date"], (e.get("url") or e["evt"]).lower()): e for e in events
            }
            for event in official:
                if id(event) in consumed:
                    continue
                merged.setdefault((event["date"], (event.get("url") or event["evt"]).lower()), event)
            events = sorted(merged.values(), key=lambda e: (e["date"], e["evt"]))
            block_years = _year_span(events, years)
    year_to_block = {year: i for i, year in enumerate(block_years)}
    start = date(min(block_years), 1, 1)
    end = date(max(block_years), 12, 31)

    prices: Dict[date, float] = {}
    dod: Dict[date, float] = {}
    trading_dod: Dict[date, float] = {}
    market_moves: Dict[str, Dict[date, float]] = {}
    if ticker:
        closes = _download_closes(ticker, start, end)
        prices, dod, trading_dod = _calendar_price_series(closes, start, end)
        for symbol in ("SPY", "XBI"):
            _, _, moves = _calendar_price_series(_download_closes(symbol, start, end), start, end)
            market_moves[symbol] = moves
    if include_price_moves and trading_dod:
        events = _add_large_move_events(events, trading_dod, move_threshold, market_moves, ticker)
    _write_enriched_events(ticker, events)
    display_events = _events_by_date(events)

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")
        rels_path = _sheet_rels_path(sheet_zip)
        rels_xml = zf.read(rels_path).decode("utf-8") if rels_path in zf.namelist() else None

    xml = _hide_legacy_columns(xml)

    # Repurpose the four blocks to the requested rolling years.
    for i, year in enumerate(block_years):
        header, legacy_anchor, date_col, price_col, dod_col, evt_col, cat_col = BLOCKS[i]
        xml = _patch_text_cell(xml, header, f"FA {year}")
        legacy_col = re.sub(r"\d", "", legacy_anchor)
        legacy_price_col = get_column_letter(column_index_from_string(legacy_col) + 1)
        xml = _patch_text_cell(xml, f"{legacy_col}8", "Date")
        xml = _patch_text_cell(xml, f"{legacy_price_col}8", "Share Price")
        xml = _patch_text_cell(xml, f"{date_col}8", "Date")
        xml = _patch_text_cell(xml, f"{price_col}8", "Share Price")
        xml = _patch_text_cell(xml, f"{dod_col}8", "DoD Chg")
        xml = _patch_text_cell(xml, f"{evt_col}8", "EVT")
        xml = _patch_text_cell(xml, f"{cat_col}8", "Category")
        xml = _patch_number_cell(xml, legacy_anchor, to_excel(datetime(year, 1, 1)))
        days = 366 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 365
        for offset in range(days):
            row = 9 + offset
            d = date(year, 1, 1) + timedelta(days=offset)
            xml = _patch_number_cell(xml, f"{date_col}{row}", to_excel(datetime(d.year, d.month, d.day)))
            if d in prices:
                xml = _patch_number_cell(xml, f"{price_col}{row}", round(prices[d], 6))
                xml = _patch_number_cell(xml, f"{dod_col}{row}", round(dod.get(d, 0.0), 10))
            else:
                xml = _patch_text_cell(xml, f"{price_col}{row}", "")
                xml = _patch_text_cell(xml, f"{dod_col}{row}", "")
            xml = _patch_text_cell(xml, f"{evt_col}{row}", "")
            xml = _patch_text_cell(xml, f"{cat_col}{row}", "")

    written = 0
    links: Dict[str, str] = {}
    for event in display_events:
        d = _parse_date(event.get("date"))
        if not d or d.year not in year_to_block:
            continue
        evt = (event.get("evt") or "").strip()
        cat = (event.get("category") or "").strip()
        if not evt:
            continue
        if cat not in VALID_CATEGORIES:
            cat = "Other"
        block_idx = year_to_block[d.year]
        _, _, _, _, _, evt_col, cat_col = BLOCKS[block_idx]
        row = 9 + (d - date(d.year, 1, 1)).days
        before = xml
        xml = _patch_text_cell(xml, f"{evt_col}{row}", evt)
        xml = _patch_text_cell(xml, f"{cat_col}{row}", cat)
        if event.get("url"):
            links[f"{evt_col}{row}"] = event["url"]
        if xml != before:
            written += 1

    xml, rels_xml = _patch_hyperlinks(xml, rels_xml, links)

    modified = {
        sheet_zip: xml.encode("utf-8"),
        rels_path: rels_xml.encode("utf-8"),
    }
    _patch_calc_mode(modified, xlsx_path)

    backup = xlsx_path.with_name(
        f"{xlsx_path.stem}_pre_events_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    shutil.copy2(xlsx_path, backup)
    tmp = xlsx_path.with_suffix(".~events.xlsx")
    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        written_parts: set[str] = set()
        for item in zin.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            if item.filename in modified:
                zout.writestr(item, modified[item.filename])
                written_parts.add(item.filename)
            else:
                zout.writestr(item, zin.read(item.filename))
        for name, payload in modified.items():
            if name not in written_parts:
                zout.writestr(name, payload)
    tmp.replace(xlsx_path)
    print(
        f"Historical Events: wrote {written} events and "
        f"{len(prices)} priced calendar rows across {block_years} → {xlsx_path}"
    )
    return written


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ticker", help="Ticker; used to auto-locate DCF file")
    ap.add_argument("--path", help="Path to DCF workbook")
    ap.add_argument("--events-file", help="JSON/CSV/markdown table of approved events")
    ap.add_argument("--years", nargs="+", type=int, help="Four block years, e.g. 2023 2024 2025 2026")
    ap.add_argument("--no-price-moves", action="store_true",
                    help="Do not add >=threshold share-price move events")
    ap.add_argument("--no-official-news", action="store_true",
                    help="Do not merge supported company official news pages")
    ap.add_argument("--require-official-news", action="store_true",
                    help="Fail closed if the official archive cannot be resolved/crawled")
    ap.add_argument("--news-url",
                    help="Company official news/press-release page URL to scrape "
                         "(official-news enrichment is skipped when omitted)")
    ap.add_argument("--company-name",
                    help="Company name; stripped from the start of official-news EVT titles")
    ap.add_argument("--news-date-selector",
                    help="Optional CSS selector for the news date on detail pages; "
                         "falls back to page meta tags when omitted")
    ap.add_argument("--news-link-pattern",
                    help="Optional regex for article hrefs on the news page; when "
                         "omitted, same-site links under the news-index path (plus "
                         "the Q4 detail path) are accepted and validated by date")
    ap.add_argument("--news-page-template",
                    help="Optional pagination template appended to the news URL for "
                         "page N (default '?page={page}')")
    ap.add_argument("--move-threshold", type=float, default=0.08,
                    help="Absolute daily share-price move threshold for auto events")
    args = ap.parse_args()

    if args.path:
        xlsx = Path(args.path)
    elif args.ticker:
        xlsx = Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/DCF {args.ticker}.xlsx")
    else:
        raise SystemExit("--path or --ticker required")
    fill_historical_events(
        xlsx,
        Path(args.events_file) if args.events_file else None,
        ticker=args.ticker,
        years=args.years,
        include_price_moves=not args.no_price_moves,
        move_threshold=args.move_threshold,
        include_official_news=not args.no_official_news,
        news_url=args.news_url,
        company_name=args.company_name,
        news_date_selector=args.news_date_selector,
        news_link_pattern=args.news_link_pattern,
        news_page_template=args.news_page_template,
        require_official_news=args.require_official_news,
    )


if __name__ == "__main__":
    main()
