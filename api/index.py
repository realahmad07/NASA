import os
import sys
import traceback

# Ensure project root is in sys.path so server.py and modules can be imported
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# Ensure writable cache directories in serverless environments
os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib")
os.environ.setdefault("NUMBA_CACHE_DIR", "/tmp/numba")

try:
    from server import app
except BaseException as e:
    err_trace = traceback.format_exc()
    try:
        from fastapi import FastAPI
        from fastapi.responses import HTMLResponse, JSONResponse

        app = FastAPI(title="PyroSage - Diagnostic Mode")

        @app.get("/api/health")
        def health_err():
            return JSONResponse(
                {"status": "degraded", "error": str(e), "trace": err_trace},
                status_code=500,
            )

        @app.get("/")
        def root_err():
            index_path = os.path.join(ROOT_DIR, "index.html")
            if os.path.exists(index_path):
                with open(index_path, "r", encoding="utf-8") as f:
                    return HTMLResponse(f.read())
            return HTMLResponse(f"<h1>PyroSage Fallback</h1><pre>{err_trace}</pre>")
    except BaseException:
        # Zero-dependency raw ASGI application fallback
        async def app(scope, receive, send):
            if scope.get("type") == "http":
                index_path = os.path.join(ROOT_DIR, "index.html")
                if os.path.exists(index_path):
                    with open(index_path, "rb") as f:
                        body = f.read()
                    status = 200
                    content_type = b"text/html; charset=utf-8"
                else:
                    body = f"<h1>PyroSage Initialization Error</h1><pre>{err_trace}</pre>".encode("utf-8")
                    status = 500
                    content_type = b"text/html; charset=utf-8"

                await send({
                    "type": "http.response.start",
                    "status": status,
                    "headers": [
                        [b"content-type", content_type],
                        [b"content-length", str(len(body)).encode("ascii")],
                    ],
                })
                await send({
                    "type": "http.response.body",
                    "body": body,
                })
