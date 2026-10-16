"""Serve the CVPR 2024 spreadsheet from one process on $PORT."""

from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "repo" / "cvpr2024_papers.xlsx"
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="CVPR 2024 Open Access")


def dataset_ready() -> bool:
    return XLSX.is_file()


@app.get("/health")
def health() -> dict[str, object]:
    return {"status": "ok", "dataset_ready": dataset_ready()}


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    if dataset_ready():
        size_mb = XLSX.stat().st_size / (1024 * 1024)
        body = (
            "<p>The CVPR 2024 Open Access spreadsheet is ready.</p>"
            f"<p><a href=\"/cvpr2024_papers.xlsx\">Download cvpr2024_papers.xlsx</a> "
            f"({size_mb:.1f} MB)</p>"
        )
    else:
        body = "<p>The spreadsheet has not been built yet. Run <code>python scrape_cvpr2024.py</code>.</p>"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>CVPR 2024 Open Access</title>
  <style>
    body {{ font-family: Georgia, serif; max-width: 40rem; margin: 4rem auto; padding: 0 1.25rem; line-height: 1.5; color: #1c1917; }}
    a {{ color: #9a3412; }}
  </style>
</head>
<body>
  <h1>CVPR 2024 Open Access</h1>
  {body}
  <p>Titles, authors, abstracts, and paper links collected from the CVF open-access listing.</p>
</body>
</html>
"""


@app.get("/cvpr2024_papers.xlsx")
def download_workbook() -> FileResponse:
    if not dataset_ready():
        raise HTTPException(status_code=404, detail="dataset not built")
    return FileResponse(XLSX, media_type=XLSX_MEDIA, filename="cvpr2024_papers.xlsx")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
