"""CVPR 2024 paper browser: refresh from the open-access site into SQLite."""

from __future__ import annotations

import html
import os
from contextlib import asynccontextmanager
from urllib.parse import quote

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

import paper_db
from scrape_cvpr2024 import LIST_URL, complete_paper, fetch_text, make_session, parse_listing

ROOT = paper_db.ROOT
XLSX = paper_db.XLSX_PATH
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@asynccontextmanager
async def lifespan(_app: FastAPI):
    paper_db.seed_from_xlsx_if_empty()
    yield


app = FastAPI(title="CVPR 2024 Open Access", lifespan=lifespan)


def dataset_ready() -> bool:
    return paper_db.count_papers() > 0 or XLSX.is_file()


@app.get("/health")
def health() -> dict[str, object]:
    return {"status": "ok", "dataset_ready": dataset_ready(), "papers": paper_db.count_papers()}


@app.post("/refresh")
def refresh_papers() -> RedirectResponse:
    try:
        papers = parse_listing(fetch_text(LIST_URL, make_session()))
    except Exception as exc:
        return RedirectResponse(f"/?error={quote(str(exc)[:180])}", status_code=303)
    if not papers:
        return RedirectResponse("/?error=" + quote("网站没有返回论文"), status_code=303)
    total = paper_db.upsert_papers(papers, keep_existing_abstract=True)
    missing = paper_db.papers_missing_abstract()
    # The listing page has no abstracts. Fill only the rows still empty so a
    # refresh stays fast when the database already holds the program.
    if len(missing) <= 40:
        for paper in missing:
            original_url = paper["paper_url"]
            try:
                complete_paper(paper)
            except Exception:
                continue
            paper_db.save_fetched_paper(original_url, paper)
    return RedirectResponse(f"/?updated={total}", status_code=303)


@app.get("/", response_class=HTMLResponse)
def index(offset: int = 0, updated: int | None = None, error: str = "") -> str:
    total = paper_db.count_papers()
    page_size = paper_db.PAGE_SIZE
    max_offset = ((total - 1) // page_size) * page_size if total else 0
    requested = max(offset, 0)
    at_end = total > 0 and requested > max_offset
    offset = min(requested, max_offset)
    rows = paper_db.list_papers(offset)
    start = offset + 1 if rows else 0
    end = offset + len(rows)
    next_offset = offset + page_size

    notices: list[str] = []
    if updated is not None:
        notices.append("已读完")
    if error:
        notices.append(error)
    if at_end:
        notices.append("已经是最后一页。")

    notice_html = "".join(f"<p class='notice'>{html.escape(item)}</p>" for item in notices)
    body_rows = []
    for row in rows:
        abstract = row["abstract"]
        preview = abstract if len(abstract) <= 180 else abstract[:180] + "…"
        title = html.escape(row["title"])
        paper_url = html.escape(row["paper_url"], quote=True)
        body_rows.append(
            "<tr>"
            f"<td>{row['position']}</td>"
            f"<td><a href=\"{paper_url}\">{title}</a></td>"
            f"<td>{html.escape(row['authors'])}</td>"
            f"<td>{html.escape(preview)}</td>"
            f"<td>{html.escape(row['pages'])}</td>"
            "</tr>"
        )
    table = (
        "<p>数据库里还没有论文。点击「读取论文」从 CVPR 2024 开放获取页面写入。</p>"
        if not rows
        else """
        <table>
          <thead>
            <tr><th>#</th><th>标题</th><th>作者</th><th>摘要</th><th>页码</th></tr>
          </thead>
          <tbody>
            {rows}
          </tbody>
        </table>
        """.format(rows="".join(body_rows))
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>CVPR 2024 论文</title>
  <style>
    body {{ font-family: Georgia, "Songti SC", serif; margin: 2rem auto; max-width: 68rem; padding: 0 1.25rem; color: #1c1917; line-height: 1.45; }}
    h1 {{ margin-bottom: 0.25rem; }}
    .count {{ font-size: 1.15rem; }}
    form {{ display: inline; }}
    button {{ font: inherit; background: #9a3412; color: white; border: 0; border-radius: 6px; padding: 0.45rem 0.9rem; cursor: pointer; margin: 0.4rem 0.5rem 0.4rem 0; }}
    button.secondary {{ background: #44403c; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 1rem; }}
    th, td {{ border-bottom: 1px solid #e7e5e4; text-align: left; vertical-align: top; padding: 0.55rem 0.45rem; }}
    th {{ font-size: 0.85rem; letter-spacing: 0.02em; }}
    td:nth-child(3), td:nth-child(4) {{ font-size: 0.92rem; }}
    a {{ color: #9a3412; }}
    .notice {{ background: #fff7ed; border: 1px solid #fed7aa; padding: 0.5rem 0.75rem; }}
    #read-status[hidden] {{ display: none; }}
  </style>
</head>
<body>
  <h1>CVPR 2024 论文</h1>
  <p class="count">论文总数：{total}</p>
  <p>当前显示第 {start}–{end} 篇。来源：<a href="{html.escape(LIST_URL, quote=True)}">CVPR 2024 Open Access</a></p>
  <p id="read-status" class="notice" hidden>正在读取。。。</p>
  {notice_html}
  <form id="read-form" method="post" action="/refresh">
    <button id="read-button" type="submit">读取论文</button>
  </form>
  <form method="get" action="/">
    <input type="hidden" name="offset" value="{next_offset}">
    <button class="secondary" type="submit">下一页</button>
  </form>
  {table}
  <script>
    document.getElementById("read-form").addEventListener("submit", function () {{
      const status = document.getElementById("read-status");
      status.hidden = false;
      status.textContent = "正在读取。。。";
      document.getElementById("read-button").disabled = true;
    }});
  </script>
</body>
</html>
"""


@app.get("/cvpr2024_papers.xlsx")
def download_workbook() -> FileResponse:
    if not XLSX.is_file():
        raise HTTPException(status_code=404, detail="dataset not built")
    return FileResponse(XLSX, media_type=XLSX_MEDIA, filename="cvpr2024_papers.xlsx")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
