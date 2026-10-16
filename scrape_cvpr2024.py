#!/usr/bin/env python3
"""Scrape CVPR 2024 Open Access metadata and write an Excel file.

The listing page has titles, authors, and links. Abstracts live on each
paper HTML page, so this script fetches the listing once, then the paper
pages with a small thread pool.
"""

from __future__ import annotations

import json
import re
import threading
import time
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote, urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://openaccess.thecvf.com"
LIST_URL = f"{BASE}/CVPR2024?day=all"
ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "repo"
OUT_PATH = OUT_DIR / "cvpr2024_papers.xlsx"
CHECKPOINT_PATH = OUT_DIR / ".cvpr2024_checkpoint.jsonl"

MAX_WORKERS = 6
MIN_INTERVAL_S = 0.08
TIMEOUT_S = 40

HEADERS = {
    "User-Agent": "cvpr2024-metadata-scraper/1.0 (personal research; +https://openaccess.thecvf.com/CVPR2024)",
    "Accept": "text/html,application/xhtml+xml",
}

ILLEGAL_XLSX = re.compile(r"[\000-\010\013\014\016-\037]")
PAGES_RE = re.compile(r"pages\s*=\s*\{([^}]*)\}", re.IGNORECASE)

COLUMNS = [
    ("title", "Title", 55),
    ("authors", "Authors", 45),
    ("abstract", "Abstract", 80),
    ("paper_url", "Paper URL", 55),
    ("pdf_url", "PDF URL", 55),
    ("supp_url", "Supplemental URL", 55),
    ("arxiv_url", "arXiv URL", 40),
    ("pages", "Pages", 16),
    ("abstract_source", "Abstract Source", 18),
]
URL_FIELDS = {"paper_url", "pdf_url", "supp_url", "arxiv_url"}

ATOM = "{http://www.w3.org/2005/Atom}"
ARXIV_ID_RE = re.compile(r"arxiv\.org/(?:abs|pdf)/(\d{4}\.\d{4,5})(?:v\d+)?")

_rate_lock = threading.Lock()
_checkpoint_lock = threading.Lock()
_arxiv_lock = threading.Lock()
_next_request_at = 0.0


class NotFound(RuntimeError):
    pass


def rate_limit() -> None:
    global _next_request_at
    with _rate_lock:
        now = time.monotonic()
        wait = _next_request_at - now
        if wait > 0:
            time.sleep(wait)
            now = time.monotonic()
        _next_request_at = now + MIN_INTERVAL_S


def make_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(HEADERS)
    return session


_thread_local = threading.local()


def thread_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = make_session()
        _thread_local.session = session
    return session


def fetch_text(url: str, session: requests.Session | None = None) -> str:
    session = session or thread_session()
    last_error: Exception | None = None
    for attempt in range(5):
        rate_limit()
        try:
            response = session.get(url, timeout=TIMEOUT_S)
            if response.status_code == 404:
                raise NotFound(url)
            if response.status_code in {429, 500, 502, 503, 504}:
                time.sleep(1.5 * (attempt + 1))
                last_error = requests.HTTPError(f"{response.status_code} for {url}")
                continue
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except requests.RequestException as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"failed to fetch {url}: {last_error}")


def clean_text(value: str) -> str:
    text = " ".join(value.replace("\xa0", " ").split())
    return ILLEGAL_XLSX.sub("", text)


def parse_listing(html: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    papers: list[dict[str, str]] = []
    for dt in soup.select("dt.ptitle"):
        title_link = dt.find("a", href=True)
        if title_link is None:
            continue
        authors_dd = dt.find_next_sibling("dd")
        links_dd = authors_dd.find_next_sibling("dd") if authors_dd else None
        if authors_dd is None or links_dd is None:
            raise RuntimeError(f"unexpected listing markup for {title_link.get_text(strip=True)!r}")

        authors = [
            clean_text(node["value"])
            for node in authors_dd.select('input[name="query_author"]')
            if node.get("value")
        ]
        pdf_url = ""
        supp_url = ""
        arxiv_url = ""
        for anchor in links_dd.find_all("a", href=True):
            label = anchor.get_text(strip=True)
            href = urljoin(BASE, anchor["href"])
            if label == "pdf":
                pdf_url = href
            elif label == "supp":
                supp_url = href
            elif label.lower() == "arxiv":
                arxiv_url = href

        bib = links_dd.select_one(".bibref")
        pages = ""
        if bib is not None:
            match = PAGES_RE.search(bib.get_text(" ", strip=True))
            if match:
                pages = clean_text(match.group(1))

        paper_url = urljoin(BASE, title_link["href"])
        papers.append(
            {
                "title": clean_text(title_link.get_text(" ", strip=True)),
                "authors": ", ".join(authors),
                "abstract": "",
                "abstract_source": "",
                "paper_url": paper_url,
                "listing_url": paper_url,
                "pdf_url": pdf_url,
                "supp_url": supp_url,
                "arxiv_url": arxiv_url,
                "pages": pages,
            }
        )
    return papers


class _ListingParser(HTMLParser):
    """Stream the CVF listing without building a document tree.

    The full listing is large enough to exhaust the 256 MB deployment
    instance when parsed with BeautifulSoup.
    """

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.papers: list[dict[str, str]] = []
        self.mode: str | None = None
        self.current: dict[str, str] | None = None
        self.authors: list[str] = []
        self.buf: list[str] = []
        self.in_title = False
        self.link_href = ""
        self.capture_link = False
        self.bib_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = {key: value or "" for key, value in attrs}
        classes = attr.get("class", "").split()
        if tag == "dt" and "ptitle" in classes:
            self.mode = "title"
            self.authors = []
            self.current = {
                "title": "",
                "authors": "",
                "abstract": "",
                "abstract_source": "",
                "paper_url": "",
                "listing_url": "",
                "pdf_url": "",
                "supp_url": "",
                "arxiv_url": "",
                "pages": "",
            }
            return
        if self.mode == "title" and tag == "a" and attr.get("href"):
            paper_url = urljoin(BASE, attr["href"])
            assert self.current is not None
            self.current["paper_url"] = paper_url
            self.current["listing_url"] = paper_url
            self.in_title = True
            self.buf = []
            return
        if tag == "dd" and self.mode == "after_title":
            self.mode = "authors"
            return
        if tag == "dd" and self.mode == "after_authors":
            self.mode = "links"
            return
        if self.mode == "authors" and tag == "input" and attr.get("name") == "query_author" and attr.get("value"):
            self.authors.append(clean_text(attr["value"]))
            return
        if self.mode == "links" and self.bib_depth == 0 and tag == "a" and attr.get("href"):
            self.link_href = urljoin(BASE, attr["href"])
            self.capture_link = True
            self.buf = []
            return
        if self.mode == "links" and tag == "div" and "bibref" in classes:
            self.bib_depth = 1
            self.buf = []
            return
        if self.bib_depth and tag == "div":
            self.bib_depth += 1

    def handle_endtag(self, tag: str) -> None:
        if tag == "a" and self.in_title and self.current is not None:
            self.current["title"] = clean_text("".join(self.buf))
            self.in_title = False
            return
        if tag == "dt" and self.mode == "title":
            self.mode = "after_title"
            return
        if tag == "a" and self.capture_link and self.current is not None:
            label = "".join(self.buf).strip()
            if label == "pdf":
                self.current["pdf_url"] = self.link_href
            elif label == "supp":
                self.current["supp_url"] = self.link_href
            elif label.lower() == "arxiv":
                self.current["arxiv_url"] = self.link_href
            self.capture_link = False
            return
        if tag == "div" and self.bib_depth:
            self.bib_depth -= 1
            if self.bib_depth == 0 and self.current is not None:
                match = PAGES_RE.search(" ".join("".join(self.buf).split()))
                if match:
                    self.current["pages"] = clean_text(match.group(1))
            return
        if tag == "dd" and self.mode == "authors":
            self.mode = "after_authors"
            return
        if tag == "dd" and self.mode == "links" and self.current is not None:
            self.current["authors"] = ", ".join(self.authors)
            if self.current["title"] and self.current["paper_url"]:
                self.papers.append(self.current)
            self.current = None
            self.mode = None

    def handle_data(self, data: str) -> None:
        if self.in_title or self.capture_link or self.bib_depth:
            self.buf.append(data)


def stream_listing(session: requests.Session) -> list[dict[str, str]]:
    last_error: Exception | None = None
    for attempt in range(3):
        parser = _ListingParser()
        response = None
        rate_limit()
        try:
            response = session.get(LIST_URL, timeout=30, stream=True)
            if response.status_code in {429, 500, 502, 503, 504}:
                last_error = requests.HTTPError(f"{response.status_code} for {LIST_URL}")
                time.sleep(1.5 * (attempt + 1))
                continue
            response.raise_for_status()
            response.encoding = "utf-8"
            for chunk in response.iter_content(chunk_size=65536, decode_unicode=True):
                if chunk:
                    parser.feed(chunk)
            parser.close()
            return parser.papers
        except requests.RequestException as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
        finally:
            if response is not None:
                response.close()
    raise RuntimeError(f"failed to fetch {LIST_URL}: {last_error}")


def parse_abstract(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    node = soup.select_one("#abstract")
    if node is None:
        return ""
    return clean_text(node.get_text(" ", strip=True))


def load_checkpoint() -> dict[str, dict]:
    found: dict[str, dict] = {}
    if not CHECKPOINT_PATH.exists():
        return found
    with CHECKPOINT_PATH.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            row = json.loads(line)
            if row.get("paper_url") and row.get("abstract"):
                found[row["paper_url"]] = row
    return found


def append_checkpoint(paper: dict[str, str]) -> None:
    record = {
        "paper_url": paper["listing_url"],
        "abstract": paper["abstract"],
        "abstract_source": paper["abstract_source"],
        "resolved_paper_url": paper["paper_url"],
        "resolved_pdf_url": paper["pdf_url"],
        "arxiv_url": paper["arxiv_url"],
    }
    line = json.dumps(record, ensure_ascii=False) + "\n"
    with _checkpoint_lock:
        with CHECKPOINT_PATH.open("a", encoding="utf-8") as handle:
            handle.write(line)


def apply_checkpoint(paper: dict[str, str], row: dict) -> None:
    paper["abstract"] = row["abstract"]
    paper["abstract_source"] = row.get("abstract_source") or "openaccess"
    if row.get("resolved_paper_url"):
        paper["paper_url"] = row["resolved_paper_url"]
    if row.get("resolved_pdf_url"):
        paper["pdf_url"] = row["resolved_pdf_url"]
    if row.get("arxiv_url"):
        paper["arxiv_url"] = row["arxiv_url"]


def url_ok(url: str) -> bool:
    rate_limit()
    response = thread_session().head(url, timeout=TIMEOUT_S, allow_redirects=True)
    return response.status_code == 200


def openaccess_abstract(paper: dict[str, str]) -> str:
    candidates = [paper["paper_url"]]
    if "paperOriginal" in paper["paper_url"]:
        candidates.append(paper["paper_url"].replace("paperOriginal", "paper"))
    for url in candidates:
        try:
            html = fetch_text(url)
        except NotFound:
            continue
        abstract = parse_abstract(html)
        if not abstract:
            continue
        paper["paper_url"] = url
        if "paperOriginal" in paper["pdf_url"]:
            replaced = paper["pdf_url"].replace("paperOriginal", "paper")
            if url_ok(replaced):
                paper["pdf_url"] = replaced
        return abstract
    return ""


def arxiv_entries(query_url: str) -> list[tuple[str, str, str]]:
    # arXiv asks for a few seconds between API calls.
    with _arxiv_lock:
        time.sleep(3)
        xml_text = fetch_text(query_url)
    root = ET.fromstring(xml_text)
    entries: list[tuple[str, str, str]] = []
    for entry in root.findall(f"{ATOM}entry"):
        title = clean_text((entry.find(f"{ATOM}title").text or "") if entry.find(f"{ATOM}title") is not None else "")
        summary_node = entry.find(f"{ATOM}summary")
        summary = clean_text(summary_node.text or "") if summary_node is not None and summary_node.text else ""
        id_node = entry.find(f"{ATOM}id")
        id_url = (id_node.text or "").strip() if id_node is not None else ""
        if title and summary and id_url:
            entries.append((title, summary, re.sub(r"v\d+$", "", id_url)))
    return entries


def arxiv_abstract(paper: dict[str, str]) -> str:
    arxiv_id = ""
    match = ARXIV_ID_RE.search(paper["arxiv_url"])
    if match:
        arxiv_id = match.group(1)
    if arxiv_id:
        entries = arxiv_entries(f"https://export.arxiv.org/api/query?id_list={arxiv_id}")
        if entries:
            return entries[0][1]

    query = 'ti:"' + paper["title"].replace('"', "") + '"'
    entries = arxiv_entries(
        "https://export.arxiv.org/api/query?search_query=" + quote(query) + "&start=0&max_results=5"
    )
    wanted = paper["title"].casefold()
    for title, summary, id_url in entries:
        if title.casefold() == wanted:
            if not paper["arxiv_url"]:
                paper["arxiv_url"] = id_url
            return summary
    return ""


def complete_paper(paper: dict[str, str]) -> None:
    abstract = openaccess_abstract(paper)
    source = "openaccess"
    if not abstract:
        abstract = arxiv_abstract(paper)
        source = "arxiv"
    if not abstract:
        raise RuntimeError(f"no abstract for {paper['title']} ({paper['listing_url']})")
    paper["abstract"] = abstract
    paper["abstract_source"] = source
    append_checkpoint(paper)


def fill_abstracts(papers: list[dict[str, str]]) -> None:
    cached = load_checkpoint()
    pending: list[int] = []
    for index, paper in enumerate(papers):
        row = cached.get(paper["listing_url"])
        if row:
            apply_checkpoint(paper, row)
        else:
            pending.append(index)
    print(f"abstracts cached={len(papers) - len(pending)} pending={len(pending)}", flush=True)
    if not pending:
        return

    done = 0
    failures: list[tuple[int, str]] = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(complete_paper, papers[index]): index for index in pending}
        for future in as_completed(futures):
            index = futures[future]
            done += 1
            try:
                future.result()
            except Exception as exc:
                failures.append((index, str(exc)))
                print(f"error {papers[index]['listing_url']}: {exc}", flush=True)
            if done % 100 == 0 or done == len(pending):
                print(f"fetched {done}/{len(pending)}", flush=True)

    if not failures:
        return

    print(f"retrying {len(failures)} papers", flush=True)
    still_missing: list[str] = []
    for index, _ in failures:
        try:
            complete_paper(papers[index])
        except Exception as exc:
            still_missing.append(f"{papers[index]['title']}: {exc}")
    if still_missing:
        raise RuntimeError(f"{len(still_missing)} abstracts still missing:\n" + "\n".join(still_missing[:10]))


def write_excel(papers: list[dict[str, str]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CVPR 2024"

    header_font = Font(bold=True)
    header_alignment = Alignment(vertical="center", wrap_text=True)
    body_alignment = Alignment(vertical="top", wrap_text=True)

    for col_index, (_, header, width) in enumerate(COLUMNS, start=1):
        cell = sheet.cell(1, col_index, header)
        cell.font = header_font
        cell.alignment = header_alignment
        sheet.column_dimensions[get_column_letter(col_index)].width = width

    for row_index, paper in enumerate(papers, start=2):
        for col_index, (field, _, _) in enumerate(COLUMNS, start=1):
            value = paper.get(field, "")
            cell = sheet.cell(row_index, col_index, value)
            cell.alignment = body_alignment
            if field in URL_FIELDS and value:
                cell.hyperlink = value
                cell.style = "Hyperlink"

    sheet.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(papers) + 1}"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.oddHeader.left.text = "CVPR 2024 Open Access"
    sheet.oddFooter.right.text = "Page &P of &N"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_TABLOID
    workbook.save(OUT_PATH)


def summarize(papers: list[dict[str, str]]) -> None:
    missing_pdf = sum(1 for paper in papers if not paper["pdf_url"])
    missing_supp = sum(1 for paper in papers if not paper["supp_url"])
    missing_arxiv = sum(1 for paper in papers if not paper["arxiv_url"])
    missing_abstract = sum(1 for paper in papers if not paper["abstract"])
    missing_authors = sum(1 for paper in papers if not paper["authors"])
    from_arxiv = sum(1 for paper in papers if paper["abstract_source"] == "arxiv")
    print(
        "papers={papers} missing_pdf={pdf} missing_supp={supp} missing_arxiv={arxiv} "
        "missing_abstract={abstract} missing_authors={authors} arxiv_abstracts={from_arxiv}".format(
            papers=len(papers),
            pdf=missing_pdf,
            supp=missing_supp,
            arxiv=missing_arxiv,
            abstract=missing_abstract,
            authors=missing_authors,
            from_arxiv=from_arxiv,
        ),
        flush=True,
    )
    if missing_pdf or missing_abstract or missing_authors:
        raise RuntimeError("listing is missing required title, author, pdf, or abstract fields")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"fetching listing {LIST_URL}", flush=True)
    papers = parse_listing(fetch_text(LIST_URL, make_session()))
    print(f"parsed {len(papers)} papers", flush=True)
    if len(papers) < 2000:
        raise RuntimeError(f"expected the full CVPR 2024 program, got {len(papers)} papers")
    urls = [paper["listing_url"] for paper in papers]
    if len(urls) != len(set(urls)):
        raise RuntimeError("duplicate paper URLs in the listing")
    fill_abstracts(papers)
    resolved = [paper["paper_url"] for paper in papers]
    if len(resolved) != len(set(resolved)):
        raise RuntimeError("duplicate resolved paper URLs")
    summarize(papers)
    write_excel(papers)
    if CHECKPOINT_PATH.exists():
        CHECKPOINT_PATH.unlink()
    print(f"wrote {OUT_PATH}", flush=True)


if __name__ == "__main__":
    main()
