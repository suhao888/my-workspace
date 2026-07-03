#!/usr/bin/env python3
"""
Excel AI 审计助手 — 本地服务
提供: 静态文件服务 + API 代理 (→ model-proxy) + CORS + 图标生成
"""

import json, os, struct, zlib, sys, io, urllib.request, urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

PORT = 3000
PROXY_URL = "http://127.0.0.1:8765"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, "src")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# ── 图标生成（纯 Python，无依赖） ──────────────────────────


def _create_png(width, height, r, g, b):
    """生成纯色 PNG 图标字节"""
    signature = b"\x89PNG\r\n\x1a\n"

    def chunk(chunk_type, data):
        c = chunk_type + data
        return (
            struct.pack(">I", len(data))
            + c
            + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        )

    # IHDR
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)

    # IDAT: 每行一个 filter byte (0) + RGB 像素
    raw = b""
    for y in range(height):
        raw += b"\x00"  # filter none
        for x in range(width):
            raw += bytes([r, g, b])

    compressed = zlib.compress(raw)
    return (
        signature
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", compressed)
        + chunk(b"IEND", b"")
    )


def ensure_icons():
    for size, name in [(16, "icon-16.png"), (32, "icon-32.png"), (80, "icon-80.png")]:
        path = os.path.join(ASSETS_DIR, name)
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(_create_png(size, size, 26, 82, 118))  # #1a5276 深蓝
            print(f"  [icon] 已生成 {name}")


# ── HTTP 服务 ──────────────────────────────────────────────

MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".json": "application/json",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"  [{self.command}] {self.path}")

    # ── CORS ──
    def _set_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")

    # ── 静态文件 ──
    def _serve_file(self, filepath):
        ext = os.path.splitext(filepath)[1].lower()
        content_type = MIME_TYPES.get(ext, "application/octet-stream")
        try:
            with open(filepath, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-cache")
            self._set_cors()
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self.send_error(404, f"File not found: {self.path}")

    def _resolve_path(self, url_path):
        if url_path == "/" or url_path == "":
            url_path = "/taskpane.html"
        clean = url_path.lstrip("/")
        # 先查 src/
        local = os.path.normpath(os.path.join(SRC_DIR, clean))
        if os.path.isfile(local):
            return local
        # 再查 assets/（URL 中的 assets/ 前缀对应 ASSETS_DIR 根目录）
        if clean.startswith("assets/"):
            local = os.path.normpath(os.path.join(ASSETS_DIR, clean[7:]))
        else:
            local = os.path.normpath(os.path.join(ASSETS_DIR, clean))
        if os.path.isfile(local):
            return local
        return None

    # ── API 代理 ──
    def _proxy_to_model(self, body_bytes):
        """转发到 model-proxy，返回 (response_json, status_code)"""
        try:
            req = urllib.request.Request(
                f"{PROXY_URL}/v1/messages",
                data=body_bytes,
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=300) as resp:
                resp_body = resp.read()
                return json.loads(resp_body), resp.status
        except urllib.error.HTTPError as e:
            try:
                return json.loads(e.read()), e.code
            except Exception:
                return {"error": f"HTTP {e.code}: {str(e)}"}, e.code
        except urllib.error.URLError as e:
            return {
                "error": f"无法连接到 AI 服务 (127.0.0.1:8765)。请确保 model-proxy 已启动。"
            }, 502
        except Exception as e:
            return {"error": f"请求失败: {str(e)}"}, 500

    # ── 路由处理 ──
    def do_OPTIONS(self):
        self.send_response(200)
        self._set_cors()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self._set_cors()
            self.end_headers()
            self.wfile.write(
                json.dumps({"status": "ok", "service": "excel-ai-addin"}).encode()
            )
            return
        filepath = self._resolve_path(parsed.path)
        if filepath:
            self._serve_file(filepath)
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/chat":
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length) if content_length > 0 else b"{}"

            resp_data, status = self._proxy_to_model(body)

            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self._set_cors()
            self.end_headers()
            self.wfile.write(json.dumps(resp_data, ensure_ascii=False).encode())
        elif parsed.path == "/api/check":
            # 检查 model-proxy 是否运行
            try:
                req = urllib.request.Request(f"{PROXY_URL}/health", method="GET")
                with urllib.request.urlopen(req, timeout=3) as resp:
                    health = json.loads(resp.read())
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self._set_cors()
                self.end_headers()
                self.wfile.write(json.dumps({"proxy": True, "health": health}).encode())
            except Exception:
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self._set_cors()
                self.end_headers()
                self.wfile.write(json.dumps({"proxy": False}).encode())
        else:
            self.send_error(404)


# ── 启动 ──
if __name__ == "__main__":
    ensure_icons()
    server = HTTPServer(("0.0.0.0", PORT), Handler)
    print(f"\n{'=' * 50}")
    print(f"  Excel AI 审计助手 — 本地服务")
    print(f"  http://localhost:{PORT}")
    print(f"  后端代理: {PROXY_URL}/v1/messages")
    print(f"{'=' * 50}")
    print(f"  1. 确保 model-proxy 在运行 (端口 8765)")
    print(f"  2. 启动 Excel → 加载项 → 上传清单")
    print(f"  3. 选择 manifest.xml")
    print(f"{'=' * 50}\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.server_close()
