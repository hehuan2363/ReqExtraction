#!/usr/bin/env python3
"""Minimal web UI to extract clauses from uploaded standards PDFs."""

from __future__ import annotations

import base64
import cgi
import html
import io
import json
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import List, Optional

try:
    from .extract_clauses import extract_pdf_data, write_xlsx
except ImportError:  # Fallback when executed as a script
    from extract_clauses import extract_pdf_data, write_xlsx

MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MiB


def truncate_text(value: str, limit: int = 220) -> tuple[str, bool]:
    if len(value) <= limit:
        return value, False
    return value[:limit].rstrip() + "…", True


def build_table(headers: List[str], rows: List[List[str]]) -> str:
    if not rows:
        return "<p>No clause content detected.</p>"
    header_html = "".join(f"<th>{html.escape(col)}</th>" for col in headers)
    body_cells = []
    for row in rows:
        cells = []
        for idx, cell in enumerate(row):
            cell_text = cell or ""
            if idx == len(row) - 1:  # Text column
                snippet, truncated = truncate_text(cell_text)
                snippet_html = html.escape(snippet).replace("\n", "<br>")
                if truncated:
                    full_attr = html.escape(cell_text, quote=True)
                    cells.append(
                        "<td class=\"text-cell\">"
                        f"<span>{snippet_html}</span> "
                        f"<button type=\"button\" class=\"more-btn\" data-full=\"{full_attr}\">More</button>"
                        "</td>"
                    )
                else:
                    cells.append(f"<td class=\"text-cell\">{snippet_html}</td>")
            else:
                cells.append(f"<td>{html.escape(cell_text)}</td>")
        body_cells.append("<tr>" + "".join(cells) + "</tr>")
    body_html = "".join(body_cells)
    return (
        "<div class=\"table-wrap\">"
        "<table>"
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{body_html}</tbody>"
        "</table>"
        "</div>"
    )


def render_page(
    message: Optional[str] = None,
    headers: Optional[List[str]] = None,
    rows: Optional[List[List[str]]] = None,
    json_b64: Optional[str] = None,
    excel_b64: Optional[str] = None,
    filename: Optional[str] = None,
) -> str:
    table_html = ""
    download_html = ""
    if headers and rows:
        table_html = build_table(headers, rows)
    if json_b64 and excel_b64:
        safe_name = html.escape(filename or "clauses")
        download_html = (
            "<div class=\"downloads\">"
            f"<a download=\"{safe_name}.json\" href=\"data:application/json;base64,{json_b64}\">Download JSON</a>"
            f"<a download=\"{safe_name}.xlsx\" href=\"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}\">Download Excel</a>"
            "</div>"
        )
    status_html = f"<p class=\"status\">{html.escape(message)}</p>" if message else ""

    return f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <title>Clause Extractor</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f9; color: #222; }}
    h1 {{ margin-bottom: 1rem; }}
    form {{ background: #fff; padding: 1.5rem; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 1.5rem; }}
    .status {{ margin-bottom: 1rem; color: #444; }}
    .downloads {{ display: flex; gap: 1rem; margin-bottom: 1rem; }}
    .downloads a {{ background: #005eb8; color: #fff; padding: 0.5rem 1rem; text-decoration: none; border-radius: 4px; }}
    .downloads a:hover {{ background: #004a91; }}
    .table-wrap {{ overflow-x: auto; background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
    table {{ border-collapse: collapse; width: 100%; min-width: 60rem; }}
    th, td {{ padding: 0.75rem; border-bottom: 1px solid #e0e0e0; vertical-align: top; text-align: left; }}
    th {{ background: #f0f4f8; }}
    .text-cell {{ max-width: 24rem; }}
    .text-cell span {{ display: inline-block; white-space: pre-wrap; }}
    .more-btn {{ margin-left: 0.5rem; background: #007a3d; border: none; color: #fff; padding: 0.25rem 0.75rem; border-radius: 4px; cursor: pointer; }}
    .more-btn:hover {{ background: #006030; }}
    .modal {{ position: fixed; inset: 0; background: rgba(0,0,0,0.6); display: flex; align-items: center; justify-content: center; }}
    .modal.hidden {{ display: none; }}
    .modal-content {{ background: #fff; padding: 1.5rem; max-width: 50rem; max-height: 80vh; overflow-y: auto; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.3); }}
    .modal-content header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; }}
    .modal-content button {{ background: #005eb8; color: #fff; border: none; padding: 0.4rem 0.9rem; border-radius: 4px; cursor: pointer; }}
    .modal-content button:hover {{ background: #004a91; }}
    pre {{ white-space: pre-wrap; margin: 0; font-family: inherit; }}
  </style>
</head>
<body>
  <h1>Standards Clause Extractor</h1>
  <form method=\"post\" action=\"/\" enctype=\"multipart/form-data\">
    <label for=\"pdf\">Select standards PDF:</label>
    <input type=\"file\" id=\"pdf\" name=\"pdf\" accept=\"application/pdf\" required>
    <button type=\"submit\">Extract</button>
    <p class=\"hint\">Maximum upload size: {MAX_UPLOAD_SIZE // (1024 * 1024)} MiB</p>
  </form>
  {status_html}
  {download_html}
  {table_html}
  <div id=\"modal\" class=\"modal hidden\">
    <div class=\"modal-content\">
      <header>
        <h2>Clause Text</h2>
        <button type=\"button\" id=\"modal-close\">Close</button>
      </header>
      <pre id=\"modal-text\"></pre>
    </div>
  </div>
  <script>
    (function() {{
      const modal = document.getElementById('modal');
      const modalText = document.getElementById('modal-text');
      const closeBtn = document.getElementById('modal-close');
      document.querySelectorAll('.more-btn').forEach(btn => {{
        btn.addEventListener('click', () => {{
          modalText.textContent = btn.dataset.full || '';
          modal.classList.remove('hidden');
        }});
      }});
      if (closeBtn) {{
        closeBtn.addEventListener('click', () => modal.classList.add('hidden'));
      }}
      modal.addEventListener('click', (event) => {{
        if (event.target === modal) {{
          modal.classList.add('hidden');
        }}
      }});
      document.addEventListener('keydown', (event) => {{
        if (event.key === 'Escape') {{
          modal.classList.add('hidden');
        }}
      }});
    }})();
  </script>
</body>
</html>
"""


class ClauseExtractionHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802 (HTTP naming)
        if self.path != "/":
            self.send_error(404)
            return
        content = render_page()
        self._send_html(content)

    def do_POST(self) -> None:  # noqa: N802 (HTTP naming)
        if self.path != "/":
            self.send_error(404)
            return
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length > MAX_UPLOAD_SIZE:
            self._send_html(render_page(message="Upload exceeds size limit."), status=413)
            return
        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": str(content_length),
        }
        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
        except (ValueError, OSError) as exc:
            self._send_html(render_page(message=f"Failed to parse upload: {exc}"), status=400)
            return
        if "pdf" not in form:
            self._send_html(render_page(message="No PDF file provided."), status=400)
            return
        file_item = form["pdf"]
        if not getattr(file_item, "file", None):
            self._send_html(render_page(message="Invalid file upload."), status=400)
            return
        file_data = file_item.file.read(MAX_UPLOAD_SIZE + 1)
        if len(file_data) == 0:
            self._send_html(render_page(message="Uploaded file is empty."), status=400)
            return
        if len(file_data) > MAX_UPLOAD_SIZE:
            self._send_html(render_page(message="Uploaded file exceeds size limit."), status=413)
            return

        suffix = Path(file_item.filename or "document.pdf").suffix or ".pdf"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file_data)
            temp_path = Path(tmp.name)
        try:
            clauses, rows = extract_pdf_data(temp_path)
        except Exception as exc:  # broad to surface errors to user
            temp_path.unlink(missing_ok=True)
            self._send_html(render_page(message=f"Failed to process PDF: {exc}"), status=500)
            return
        temp_path.unlink(missing_ok=True)

        json_payload = json.dumps([clause.to_dict() for clause in clauses], indent=2)
        json_b64 = base64.b64encode(json_payload.encode("utf-8")).decode("ascii")

        buffer = io.BytesIO()
        write_xlsx(rows, buffer)
        excel_b64 = base64.b64encode(buffer.getvalue()).decode("ascii")

        message = f"Extracted {len(rows) - 1} clauses from {file_item.filename or 'uploaded file'}."
        content = render_page(
            message=message,
            headers=rows[0],
            rows=rows[1:],
            json_b64=json_b64,
            excel_b64=excel_b64,
            filename=(file_item.filename or "clauses").rsplit(".", 1)[0],
        )
        self._send_html(content)

    def _send_html(self, content: str, status: int = 200) -> None:
        body = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run(host: str = "127.0.0.1", port: int = 8000) -> None:
    address = (host, port)
    httpd = HTTPServer(address, ClauseExtractionHandler)
    print(f"Serving on http://{host}:{port}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()


if __name__ == "__main__":
    run()
