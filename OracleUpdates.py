#!/usr/bin/env python3
"""
Oracle Updates Summarizer (kind of a personal helper script)
=================================================================

What this does (in plain language):
    1. Looks inside an Excel sheet (we assume the feature summary export) and grabs ANY hyperlinks
         that live in a specific column (default column C). We skip the header row because usually
         row 1 is just a label.
    2. For every link we find, we fire up a Chrome browser with Selenium (headless if you ask)
         and pull down the page. We try to strip out the readable text.
    3. We then call GPT-5 (through OpenAI Responses API) with a JSON schema to coerce a structured
         summary (title, product area, release, etc.). If the model fails or returns junk, we still
         put something in the final report so you know it failed.
    4. Finally we write a Markdown file grouping things by release > product area.

Why all the extra complexity in Excel parsing? Because Excel sometimes stores links as formulas,
sometimes as relationship hyperlinks, and occasionally people just paste the raw URL text.
openpyxl sometimes misbehaves with certain edge cases (shared strings, malformed XML), so we have
an ugly-but-robust ZIP/XML fallback method.

Usage example (not super strict; flags are optional):
    python oracle_update_summarizer.py \
            --excel "/mnt/data/Feature_Summary_8_13_2025.xlsx" \
            --sheet "Sheet1" \
            --out "oracle_update_summary_25C.md" \
            --headless

Environment:
    You MUST export an OPENAI_API_KEY or else the script will refuse to continue.

Install (rough list; if something errors, just pip install it):
    pip install openpyxl pandas beautifulsoup4 lxml selenium webdriver-manager tenacity openai tqdm

NOTE: This script is intentionally a little "chatty" in comments to make it easier for future-me
or someone new to Python to follow along. Some bits might look over-explained; that's on purpose.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass  # dataclass = convenient way to store structured data (like a mini record)
from typing import Iterable, List, Optional, Tuple

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re, zipfile, xml.etree.ElementTree as ET
from typing import List, Optional, Tuple
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from tqdm import tqdm
import contextlib

try:  # httpx used only when we need explicit proxy / custom CA handling for OpenAI
    import httpx  # type: ignore
except Exception:  # pragma: no cover - optional
    httpx = None

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- OpenAI (Responses API) ---
from openai import OpenAI
#from openai.types.chat.completion_create_params import ResponseFormat

# In a container context we'll typically mount a host folder to /data. Provide a saner default
# (the user can override with --excel). Leaving the original example in the top comment.
DEFAULT_EXCEL = "/data/Feature_Summary.xlsx"


@dataclass
class PageSummary:
    """Simple container for the structured GPT output.

    (Could have used a plain dict, but typed attributes make it a little clearer.)
    """
    url: str               # the original page URL
    title: str             # feature / change title
    product_area: Optional[str]
    release: Optional[str]   # e.g., "25C" (Oracle quarterly naming pattern)
    change_type: Optional[str]  # new feature / change / fix / etc.
    summary: str           # the model's summary write-up
    business_impact: Optional[str]  # how it matters to customers / org
    actions: Optional[str]          # any setup / enablement / testing instructions
    tags: List[str]        # loose tag list for filtering later


def _setup_logger(verbosity: int) -> None:
    """Basic logger setup.

    I keep this really small; for a "serious" project I'd maybe configure handlers.
    Verbosity levels:
      0 -> WARNING (default)
      1 -> INFO
      2+-> DEBUG
    """
    level = logging.WARNING
    if verbosity == 1:
        level = logging.INFO
    elif verbosity >= 2:
        level = logging.DEBUG
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


def extract_links_from_excel(
    excel_path: str,
    sheet_name: Optional[str] = None,
    column_letter: str = "C",
    start_row: int = 2,
) -> List[Tuple[str, Optional[str]]]:
    """Pull all URL-ish things out of a single Excel column.

    We deliberately over-engineer extraction because real-world Excel files can be messy.

    Strategy (fast then fallback):
      1. Use openpyxl in read_only mode (fast and memory-friendly) to iterate just that column.
      2. If that blows up (rare edge cases), dive into the XLSX zip structure ourselves and read
         the underlying XML (relationship hyperlinks + raw HYPERLINK formulas + plain text).

    Args:
      excel_path: path to the .xlsx file.
      sheet_name: name of worksheet (if None we just take the active first sheet).
      column_letter: Excel column letter to parse (default 'C').
      start_row: skip header rows before this (default 2 means row 1 is header).

    Returns:
      List of tuples (url, display_text_or_None)
    """
    try:
        # 1) FAST PATH: streaming read just column C in the requested sheet
        wb = load_workbook(excel_path, data_only=False, read_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active

        col_idx = column_index_from_string(column_letter)
        out: List[Tuple[str, Optional[str]]] = []

        for (cell,) in ws.iter_rows(min_row=start_row, min_col=col_idx, max_col=col_idx):
            if cell.value in (None, ""):
                continue

            url = None
            text = None

            # True hyperlink object (if present)
            if getattr(cell, "hyperlink", None) and getattr(cell.hyperlink, "target", None):
                url = cell.hyperlink.target
                text = str(cell.value) if cell.value is not None else None
            else:
                # =HYPERLINK("url","text") formula or raw URL string
                val = str(cell.value)
                m = re.match(r'^\s*=\s*HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)\s*$', val, re.IGNORECASE)
                if m:
                    url, text = m.group(1), m.group(2)
                elif re.match(r"^https?://", val, re.IGNORECASE):
                    url, text = val.strip(), None

            if url:
                out.append((url, text))

        return out  # success path
    except Exception as e:  # noqa: F841 (broad for resilience; fallback below)
        return _extract_links_via_zip(excel_path, sheet_name, column_letter, start_row)


def _extract_links_via_zip(
    excel_path: str,
    sheet_name: Optional[str],
    column_letter: str,
    start_row: int
) -> List[Tuple[str, Optional[str]]]:
    """Low-level Excel hyperlink extraction ("DIY mode").

    Plan B when openpyxl doesn't cooperate. XLSX files are just ZIP archives of XML files.
    """
    # openpyxl really did not like their excel for some reason. Struggled for a little bit
    col_letter = column_letter.upper()

    with zipfile.ZipFile(excel_path, "r") as z:
        # Map r:id -> Target for workbook
        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        nsr = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        wb_rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels.findall(".//{*}Relationship")}

        # Find the sheet record in workbook.xml
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        # default namespace dance
        ns = {"s": wb.tag.split("}")[0].strip("{")}

        target_sheet = None
        for sh in wb.findall(".//s:sheets/s:sheet", ns):
            nm = sh.attrib.get("name")
            if (sheet_name and nm == sheet_name) or (not sheet_name and True):
                # If no sheet_name provided, first sheet wins
                if sheet_name is None or nm == sheet_name:
                    rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    sheet_rel_target = wb_rel_map[rid]  # e.g., 'worksheets/sheet1.xml'
                    target_sheet = f"xl/{sheet_rel_target}"
                    if sheet_name is None:
                        # if we took the first sheet without name, stop after first
                        break
                    else:
                        # matched by explicit name, stop
                        break

        if not target_sheet:
            raise RuntimeError(f"Sheet not found: {sheet_name!r}")

        # Parse the sheet XML
        sheet_xml = ET.fromstring(z.read(target_sheet))
        sheet_dir = "/".join(target_sheet.split("/")[:-1])
        # Load sheet rels if present
        rels_path = f"{sheet_dir}/_rels/{target_sheet.split('/')[-1]}.rels"
        rels_map = {}
        if rels_path in z.namelist():
            srels = ET.fromstring(z.read(rels_path))
            rels_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in srels.findall(".//{*}Relationship")}

        # Collect relationship-style hyperlinks, keyed by cell ref (A1 notation)
        hyperlink_targets = {}
        for h in sheet_xml.findall(".//{*}hyperlinks/{*}hyperlink"):
            ref = h.attrib.get("ref")
            rid = h.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if ref and rid and rid in rels_map:
                hyperlink_targets[ref] = rels_map[rid]

        # Iterate rows >= start_row; pick only cells whose ref starts with the desired column (e.g., 'C')
        links: List[Tuple[str, Optional[str]]] = []
        for row in sheet_xml.findall(".//{*}sheetData/{*}row"):
            r_idx = int(row.attrib.get("r", "0") or "0")
            if r_idx < start_row:
                continue
            for c in row.findall("{*}c"):
                cell_ref = c.attrib.get("r", "")  # e.g., "C2"
                if not cell_ref.startswith(col_letter):
                    continue

                # 1) Relationship-style hyperlink?
                if cell_ref in hyperlink_targets:
                    url = hyperlink_targets[cell_ref]
                    # Optional display text: read string value if present
                    text = None
                    # If the cell has inline string (t="inlineStr") or shared string, try to read it
                    t = c.attrib.get("t")
                    if t == "inlineStr":
                        is_node = c.find("{*}is")
                        if is_node is not None:
                            t_node = is_node.find("{*}t")
                            if t_node is not None and t_node.text:
                                text = t_node.text
                    elif t == "s":
                        # sharedStrings lookup (best-effort; if missing, leave None)
                        try:
                            ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
                            v_node = c.find("{*}v")
                            if v_node is not None and v_node.text is not None:
                                idx = int(v_node.text)
                                si = ss.findall("{*}si")[idx]
                                t_node = si.find("{*}t")
                                if t_node is not None and t_node.text:
                                    text = t_node.text
                        except KeyError:
                            pass
                    links.append((url, text))
                    continue

                # 2) Formula HYPERLINK("url","text")?
                f_node = c.find("{*}f")
                if f_node is not None and f_node.text:
                    m = re.search(r'HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)', f_node.text, re.IGNORECASE)
                    if m:
                        url, text = m.group(1), m.group(2)
                        links.append((url, text))
                        continue

                # 3) Raw URL in text?
                v_node = c.find("{*}v")
                if v_node is not None and v_node.text:
                    # Might be shared string index
                    t = c.attrib.get("t")
                    if t == "s":
                        # Resolve shared string
                        try:
                            ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
                            idx = int(v_node.text)
                            si = ss.findall("{*}si")[idx]
                            t_node = si.find("{*}t")
                            raw = t_node.text if (t_node is not None and t_node.text) else ""
                        except KeyError:
                            raw = v_node.text or ""
                    else:
                        raw = v_node.text or ""
                    raw = str(raw).strip()
                    if re.match(r"^https?://", raw, re.IGNORECASE):
                        links.append((raw, None))

        return links


def make_driver(headless: bool = True, proxy: Optional[str] = None, insecure: bool = False) -> webdriver.Chrome:
    """Spin up a Selenium Chrome driver (headless optional, with optional proxy / insecure flags).

    Args:
        headless: run Chrome in headless mode.
        proxy: proxy URL like http://user:pass@host:port or socks5://host:1080 .
        insecure: if True, ignore certificate errors (NOT recommended for production; last resort).
    """
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1366,1024")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-allow-origins=*")
    options.add_argument("--lang=en-US")
    if proxy:
        options.add_argument(f"--proxy-server={proxy}")
    if insecure:
        # Corporate TLS inspection sometimes breaks Chrome if root CA not installed; this bypasses verification.
        # Prefer adding the custom root CA instead of using this.
        options.add_argument("--ignore-certificate-errors")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    driver.set_script_timeout(60)
    return driver


def extract_main_text_from_html(html: str) -> Tuple[str, str]:
    """Very basic content extraction (title + main-ish text)."""
    soup = BeautifulSoup(html, "lxml")

    # Title
    title_tag = soup.find("title")
    title = (title_tag.get_text(strip=True) if title_tag else "").strip()

    # Remove script/style/nav/footer noise
    for bad in soup(["script", "style", "noscript"]):
        bad.decompose()

    # Oracle readiness pages often have a main article/section; fall back to body
    candidates = []
    for sel in ["article", "main", "section", "div#main", "div.content", "div#content"]:
        node = soup.select_one(sel)
        if node and node.get_text(strip=True):
            candidates.append(node)

    node = candidates[0] if candidates else soup.body or soup
    text = "\n".join(
        line for line in (node.get_text("\n", strip=True)).splitlines() if line.strip()
    )

    return title, text


@retry(
    stop=stop_after_attempt(2),
    wait=wait_exponential(multiplier=1, min=1, max=6),
    retry=retry_if_exception_type(Exception),
    reraise=True,
)
def fetch_page(driver: webdriver.Chrome, url: str) -> Tuple[str, str]:
    """Navigate to a URL and pull out (title, text).

    Decorated with tenacity.retry so transient hiccups (network blips, slow loads) get one retry.
    We intentionally keep waits simple: wait for <body>, then short sleep for dynamic JS.
    """
    driver.get(url)

    # Wait for body/content to be present
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )

    # Selinum gets 1 second to load, if not fast enough, it fails, and this repeats
    time.sleep(1)

    html = driver.page_source
    return extract_main_text_from_html(html)


def gpt5_client(proxy: Optional[str] = None, ca_bundle: Optional[str] = None) -> OpenAI:
    """Create an OpenAI client instance, honoring optional proxy and custom CA bundle.

    The OpenAI Python SDK already respects HTTPS_PROXY / HTTP_PROXY env vars. We only build a
    custom httpx client if we explicitly need to force a proxy or CA file. This helps inside
    locked-down networks where TLS interception causes EOF / protocol violations.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY environment variable is not set.")

    client_kwargs = {}
    if proxy or ca_bundle:
        if httpx is None:
            raise RuntimeError("httpx not available to configure proxy/CA; install httpx or remove --proxy/--ca-bundle")
        verify: Optional[object]
        if ca_bundle:
            verify = ca_bundle  # path to PEM bundle
        else:
            verify = True
        # Allow environment variables to still complement configuration.
        http_client = httpx.Client(proxies=proxy, verify=verify, timeout=60.0)
        client_kwargs["http_client"] = http_client
    return OpenAI(**client_kwargs)


# JSON schema to enforce a consistent, parseable summary structure
def summary_schema():
    """Return the JSON schema we tell GPT to obey.

    Having this factored lets us evolve fields easily (add/remove) without hunting in multiple spots.
    The model (with strict mode) should only emit these keys.
    """
    return {
        "type": "object",
        "properties": {
            "title": {"type": "string"},
            "product_area": {"type": ["string", "null"], "description": "e.g., HCM Common, Talent, Time, Absence"},
            "release": {"type": ["string", "null"], "description": "e.g., 25C"},
            "change_type": {"type": ["string", "null"], "description": "New Feature | Change | Fix | Deprecation"},
            "summary": {"type": "string"},
            "business_impact": {"type": ["string", "null"]},
            "actions_required": {"type": ["string", "null"], "description": "Setup steps, opt-in flags, testing notes"},
            "tags": {"type": "array", "items": {"type": "string"}},
            "source_url": {"type": "string"},
        },
        "required": [
            "title",
            "product_area",
            "release",
            "change_type",
            "summary",
            "business_impact",
            "actions_required",
            "tags",
            "source_url",
        ],
        "additionalProperties": False,
    }



def build_prompt(page_text: str, url: str) -> List[dict]:
    """Construct the messages list for the Responses API.

    (Could inline this, but separating makes experimenting with system vs user prompts easier.)
    """
    SYS = (
        "You are an expert Oracle Cloud Applications Readiness analyst. "
        "Extract crisp, accurate details from an Oracle Readiness article. "
        "Prefer facts in the page over assumptions. If a field is unknown, leave it blank."
        "Speak corporate language."
    )

    USER = (
        "Summarize the following Oracle Readiness page into a compact, practical brief for HCM admins. "
        "Capture the product area, release (e.g., 25C), change type, and key actions if present.\n\n"
        f"URL: {url}\n\n"
        "----- PAGE TEXT START -----\n"
        f"{page_text[:15000]}\n"  # hard cap to avoid hitting token limits; GPT-5 has large context but be safe
        "----- PAGE TEXT END -----\n"
    )

    return [
        {"role": "developer", "content": SYS},
        {"role": "user", "content": USER},
    ]


def summarize_with_gpt5(client: OpenAI, page_text: str, url: str) -> PageSummary:
    """Send the page text off to GPT-5 and coerce into PageSummary.

    If the model returns invalid JSON (happens sometimes even with strict mode), we fall back
    to a minimal structure using whatever raw text we got.
    """
    #schema = summary_schema() # (currently unused directly but could validate locally later)

    # Responses API expects a particular structure for json schema formatting.
    response_format = {
        "format": {
            "type": "json_schema",
            "name": "oracle_readiness_summary",
            "schema": summary_schema(),  # includes required keys
            "strict": True,              # enforce only those keys
        }
    }

    model = os.getenv("MODEL", "gpt-5")  # allow overriding if environment dictates

    # Fire the request. (Could add retry but already have retry on page fetch; usually stable.)
    # Had issue here where response format was just... Not working.... I was silly

    resp = client.responses.create(
        model=model,
        input=build_prompt(page_text, url),
        text=response_format,
        reasoning={"effort": "minimal"},  # keep it cheap; we don't need deep reasoning
        max_output_tokens=800,
    )

    # Grab convenience text if present; else manually walk the response blocks.
    out_text = getattr(resp, "output_text", "") or ""
    if not out_text:
        try:
            for block in getattr(resp, "output", []) or []:
                for c in getattr(block, "content", []) or []:
                    t = getattr(c, "text", None)
                    if t:
                        out_text += t
        except Exception:  # swallow—parsing fallbacks below
            pass

    # Try parse JSON; fallback to stub if invalid.
    try:
        data = json.loads(out_text)
    except Exception:
        data = {
            "title": "",
            "product_area": None,
            "release": None,
            "change_type": None,
            "summary": (out_text or "").strip()[:4000],  # truncate so we don't bloat file
            "business_impact": None,
            "actions_required": None,
            "tags": [],
            "source_url": url,
        }

    # Return a nice strongly-typed dataclass instance.
    return PageSummary(
        url=data.get("source_url", url),
        title=(data.get("title") or "").strip(),
        product_area=(data.get("product_area") or None),
        release=(data.get("release") or None),
        change_type=(data.get("change_type") or None),
        summary=(data.get("summary") or "").strip(),
        business_impact=(data.get("business_impact") or data.get("business impact") or None),
        actions=(data.get("actions_required") or None),
        tags=[t for t in (data.get("tags") or []) if isinstance(t, str)],
    )



def write_markdown(summaries: List[PageSummary], out_path: str) -> None:
    """Write a consolidated Markdown report.

    The grouping order (release > product area) is mostly for quick scanning. Nothing fancy like
    a TOC—keeping it simple so it opens nicely in any basic viewer.
    """
    def key(s: PageSummary):
        # Sorting key; using empty strings for None keeps ordering stable.
        return (s.release or "", s.product_area or "", s.title or s.url)

    ordered = sorted(summaries, key=key)

    lines: List[str] = []
    lines.append("# Oracle HCM Readiness — Newest Update Summary\n")
    lines.append(
        f"_Generated by oracle_update_summarizer.py on {pd.Timestamp.now().date().isoformat()}_\n"
    )

    current_release = None
    current_area = None

    for s in ordered:
        # New release section header
        if s.release != current_release:
            current_release = s.release
            lines.append(f"\n## {current_release or 'Unspecified Release'}\n")

        # New product area sub-section
        if s.product_area != current_area:
            current_area = s.product_area
            lines.append(f"\n### {current_area or 'General'}\n")

        title = s.title or "(Untitled)"
        lines.append(f"#### {title}")
        lines.append(f"- **Source:** {s.url}")
        if s.change_type:
            lines.append(f"- **Change Type:** {s.change_type}")
        if s.tags:
            lines.append(f"- **Tags:** {', '.join(s.tags)}")
        lines.append("")
        lines.append(s.summary or "_No summary returned._")
        if s.business_impact:
            lines.append(f"\n**Business Impact:** {s.business_impact}")
        if s.actions:
            lines.append(f"\n**Actions / Setup:** {s.actions}")
        lines.append("\n---\n")

    # Actually write the file
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    """CLI entrypoint.

    Note: This isn't packaged—just run the script directly. We keep logic inline for clarity.
    """
    parser = argparse.ArgumentParser(
        description="Summarize Oracle Readiness updates from Excel links using GPT-5."
    )
    parser.add_argument("--excel", default=DEFAULT_EXCEL, help="Path to Excel file (Feature Summary).")
    parser.add_argument("--sheet", default=None, help="Sheet name (defaults to active sheet).")
    parser.add_argument("--out", default="oracle_update_summary.md", help="Output Markdown filename.")
    parser.add_argument("--headless", action="store_true", help="Run Chrome headless.")
    parser.add_argument("--proxy", default=None, help="Proxy URL for outbound HTTP/S (e.g. http://user:pass@host:port). Overrides HTTPS_PROXY env var if set.")
    parser.add_argument("--no-proxy", default=None, help="Comma-separated hosts to bypass proxy (adds to NO_PROXY env during run).")
    parser.add_argument("--ca-bundle", default=None, help="Path to custom CA bundle PEM (mounted inside container). Sets SSL_CERT_FILE & REQUESTS_CA_BUNDLE.")
    parser.add_argument("--insecure", action="store_true", help="Ignore TLS certificate errors in Chrome ONLY (last resort; prefer --ca-bundle).")
    parser.add_argument(
        "--limit", type=int, default=None, help="Max number of links to process (debug / faster trial)."
    )
    parser.add_argument(
        "-v", "--verbose", action="count", default=0, help="Increase verbosity (repeat for more)."
    )
    args = parser.parse_args()

    _setup_logger(args.verbose)

    # Proxy / CA handling early so all later libs see env vars.
    proxy = args.proxy or os.getenv("HTTPS_PROXY") or os.getenv("https_proxy") or os.getenv("HTTP_PROXY") or os.getenv("http_proxy")
    if args.no_proxy:
        existing = os.getenv("NO_PROXY") or os.getenv("no_proxy") or ""
        merged = ",".join(filter(None, [existing, args.no_proxy]))
        os.environ["NO_PROXY"] = merged
    if args.ca_bundle:
        # Point Python / requests / httpx to custom CA
        os.environ["SSL_CERT_FILE"] = args.ca_bundle
        os.environ["REQUESTS_CA_BUNDLE"] = args.ca_bundle
    if proxy:
        # Populate standard envs so any library (selenium downloads etc.) can reuse
        os.environ.setdefault("HTTPS_PROXY", proxy)
        os.environ.setdefault("HTTP_PROXY", proxy)
        logging.info("Using proxy: %s", proxy)
    if args.ca_bundle:
        logging.info("Using custom CA bundle: %s", args.ca_bundle)

    # STEP 1: Extract links
    logging.info("Reading Excel and extracting links...")
    links = extract_links_from_excel(args.excel, sheet_name=args.sheet, column_letter="C", start_row=2)
    if not links:
        print("No links found in column C (rows >= 2).", file=sys.stderr)
        sys.exit(2)

    if args.limit:
        links = links[: args.limit]  # quick manual limit for testing

    # STEP 2: Fetch pages with Selenium
    logging.info("Launching Selenium Chrome...")
    driver = make_driver(headless=args.headless, proxy=proxy, insecure=args.insecure)

    client = gpt5_client(proxy=proxy, ca_bundle=args.ca_bundle)
    summaries: List[PageSummary] = []

    try:
        for url, text in tqdm(links, desc="Processing links", unit="page"):
            try:
                title, page_text = fetch_page(driver, url)
                merged_text = f"{title}\n\n{page_text}" if title else page_text
                summary = summarize_with_gpt5(client, merged_text, url)
                if not summary.title and title:  # fallback to DOM title if model omitted it
                    summary.title = title
                summaries.append(summary)
            except Exception as e:
                logging.exception("Failed to process %s: %s", url, e)
                # Keep a placeholder so user still sees there was a URL attempted.
                summaries.append(
                    PageSummary(
                        url=url,
                        title="(Failed to summarize)",
                        product_area=None,
                        release=None,
                        change_type=None,
                        summary=f"Error during fetch/summarize: {e}",
                        business_impact=None,
                        actions=None,
                        tags=[],
                    )
                )
                continue  # move on; best-effort philosophy
    finally:
        with contextlib.suppress(Exception):
            driver.quit()
        # Close custom http client if we created one
        http_client = getattr(client, "_client", None)  # openai 1.x stores underlying
        with contextlib.suppress(Exception):
            if http_client and hasattr(http_client, "close"):
                http_client.close()

    # STEP 3: Write report
    write_markdown(summaries, args.out)
    print(f"Done. Wrote: {args.out}  (summarized {len(summaries)} pages)")


if __name__ == "__main__":
    main()
