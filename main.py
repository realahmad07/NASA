"""PyroSage main entrypoint for Vercel and framework runners."""
import os
import sys
import traceback

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

try:
    from server import app
except BaseException as e:
    err_trace = traceback.format_exc()
    try:
        from fastapi import FastAPI
        from fastapi.responses import HTMLResponse

        app = FastAPI(title="PyroSage - Diagnostic Mode")

        @app.get("/")
        def root():
            index_path = os.path.join(ROOT_DIR, "index.html")
            if os.path.exists(index_path):
                with open(index_path, "r", encoding="utf-8") as f:
                    return HTMLResponse(f.read())
            return HTMLResponse(f"<h1>Error</h1><pre>{err_trace}</pre>")
    except BaseException:
        async def app(scope, receive, send):
            if scope.get("type") == "http":
                index_path = os.path.join(ROOT_DIR, "index.html")
                if os.path.exists(index_path):
                    with open(index_path, "rb") as f:
                        body = f.read()
                    status = 200
                else:
                    body = f"<h1>Error</h1><pre>{err_trace}</pre>".encode("utf-8")
                    status = 500

                await send({
                    "type": "http.response.start",
                    "status": status,
                    "headers": [
                        [b"content-type", b"text/html; charset=utf-8"],
                        [b"content-length", str(len(body)).encode("ascii")],
                    ],
                })
                await send({"type": "http.response.body", "body": body})

__all__ = ["app"]
