"""Local test fixture for the already configured API path; never deploy this as an API."""
from email.utils import formatdate
from hashlib import sha256
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parent
FILES = {
    "/existing-download-api/files/policy-001": (
        ROOT / "demo-files" / "star-young-writer-photography-role.docx",
        "star-young-writer-photography-role.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ),
    "/existing-download-api/files/form-002": (
        ROOT / "demo-files" / "service-record-template.xlsx",
        "service-record-template.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ),
    "/existing-download-api/files/notice-003": (
        ROOT / "demo-files" / "compliance-notice.docx",
        "compliance-notice.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ),
    "/existing-download-api/files/guide-004": (
        ROOT / "demo-files" / "cache-usage-guide.pdf",
        "cache-usage-guide.pdf",
        "application/pdf",
    ),
}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        path = unquote(urlsplit(self.path).path)
        if "demo-files" in path.split("/"):
            self.send_error(404)
            return
        entry = FILES.get(path)
        if entry is None:
            return super().do_GET()
        doc, filename, mime_type = entry
        data = doc.read_bytes()  # Python fixture only; front-end caches the original Response stream.
        etag = '"' + sha256(data).hexdigest() + '"'
        modified = formatdate(doc.stat().st_mtime, usegmt=True)
        if self.headers.get("If-None-Match") == etag or (
            not self.headers.get("If-None-Match")
            and self.headers.get("If-Modified-Since") == modified
        ):
            self.send_response(304)
            self.send_header("ETag", etag)
            self.send_header("Last-Modified", modified)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("ETag", etag)
        self.send_header("Last-Modified", modified)
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(data)


if __name__ == "__main__":
    print("Local fixture: http://127.0.0.1:8080/index.html", flush=True)
    ThreadingHTTPServer(("127.0.0.1", 8080), Handler).serve_forever()
