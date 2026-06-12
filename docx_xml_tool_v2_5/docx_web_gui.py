from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs

from docx_workspace_tool import apply_markdown_workspace, export_docx_xml, fill_fields_workspace

VERSION = "2.5"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765


class DocxWebGuiApp:
    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir).resolve()

    def handle_request(self, method: str, path: str, body: bytes, headers: dict[str, str] | None = None):
        headers = headers or {}
        try:
            if method == "GET" and path == "/":
                return html_response(render_home_page())
            if method == "GET" and path == "/api/health":
                return json_response({"status": "ok", "version": VERSION})
            if method == "POST" and path == "/api/upload-docx":
                filename, content = parse_multipart_file(body, headers.get("Content-Type", ""), "docx_file")
                if Path(filename).suffix.lower() != ".docx":
                    return json_response({"ok": False, "error": "请拖入 .docx 文件。"}, 400)
                upload_dir = self.base_dir / "gui_uploads"
                upload_dir.mkdir(parents=True, exist_ok=True)
                upload_path = upload_dir / safe_filename(filename)
                upload_path.write_bytes(content)
                workspace = export_docx_xml(upload_path, workspace_root=self.base_dir / "docx_xml_outputs")
                return json_response(workspace_payload(workspace))
            if method == "POST" and path == "/api/export":
                data = parse_form(body)
                workspace = export_docx_xml(data["input_docx"], workspace_root=data.get("workspace_root") or None)
                return json_response({"ok": True, "workspace": str(workspace)})
            if method == "POST" and path == "/api/apply-md":
                data = parse_form(body)
                output = apply_markdown_workspace(data["workspace"], data.get("output") or None)
                return json_response({"ok": True, "output": str(output)})
            if method == "POST" and path == "/api/fill-fields":
                data = parse_form(body)
                output = fill_fields_workspace(data["workspace"], data.get("field_values") or None, data.get("output") or None)
                return json_response({"ok": True, "output": str(output)})
            if method == "POST" and path == "/api/open-path":
                data = parse_form(body)
                target = resolve_gui_path(self.base_dir, data["path"])
                if data.get("test_mode") != "1":
                    open_local_path(target)
                return json_response({"ok": True, "path": str(target)})
        except KeyError as exc:
            return json_response({"ok": False, "error": f"缺少参数：{exc.args[0]}"}, 400)
        except ValueError as exc:
            return json_response({"ok": False, "error": str(exc)}, 400)
        except Exception as exc:
            return json_response({"ok": False, "error": str(exc)}, 500)
        return text_response("Not found", 404)


def create_app(base_dir: Path | str | None = None) -> DocxWebGuiApp:
    return DocxWebGuiApp(Path(base_dir) if base_dir else Path.cwd())


def parse_form(body: bytes) -> dict[str, str]:
    parsed = parse_qs(body.decode("utf-8"), keep_blank_values=True)
    return {key: values[-1].strip() for key, values in parsed.items()}


def parse_multipart_file(body: bytes, content_type: str, field_name: str) -> tuple[str, bytes]:
    match = re.search(r"boundary=([^;]+)", content_type or "")
    if not match:
        raise ValueError("缺少上传边界。")
    boundary = ("--" + match.group(1).strip().strip('"')).encode("utf-8")
    for part in body.split(boundary):
        part = part.strip(b"\r\n")
        if not part or part == b"--" or b"\r\n\r\n" not in part:
            continue
        raw_headers, content = part.split(b"\r\n\r\n", 1)
        header_text = raw_headers.decode("utf-8", errors="replace")
        if f'name="{field_name}"' not in header_text:
            continue
        filename_match = re.search(r'filename="([^"]+)"', header_text)
        if not filename_match:
            raise ValueError("上传内容缺少文件名。")
        return filename_match.group(1), content.rstrip(b"\r\n")
    raise ValueError("没有找到上传的 DOCX 文件。")


def safe_filename(filename: str) -> str:
    name = Path(filename).name or "uploaded.docx"
    return re.sub(r"[^一-鿿A-Za-z0-9._ -]+", "_", name)


def workspace_payload(workspace: Path):
    workspace = Path(workspace)
    return {
        "ok": True,
        "workspace": str(workspace),
        "document_md": str(workspace / "document.md"),
        "field_values": str(workspace / "field_values.yaml"),
        "edits_dir": str(workspace / "edits"),
    }


def resolve_gui_path(base_dir: Path, value: str) -> Path:
    base_dir = Path(base_dir).resolve()
    target = Path(value).resolve()
    if target != base_dir and base_dir not in target.parents:
        raise ValueError("只能打开当前工具工作区内的文件。")
    return target


def open_local_path(path: Path) -> None:
    path = Path(path)
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def json_response(payload: dict, status: int = 200):
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    return status, {"Content-Type": "application/json; charset=utf-8"}, body


def html_response(html: str, status: int = 200):
    return status, {"Content-Type": "text/html; charset=utf-8"}, html.encode("utf-8")


def text_response(text: str, status: int = 200):
    return status, {"Content-Type": "text/plain; charset=utf-8"}, text.encode("utf-8")


def render_home_page() -> str:
    return """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>DOCX XML Tool v2.5</title>
  <style>
    :root { color-scheme: dark; --bg: #10130f; --panel: #f3ead7; --ink: #21190f; --muted: #75644d; --accent: #c7ff45; --hot: #ff6a3d; }
    * { box-sizing: border-box; }
    body { margin: 0; min-height: 100vh; background: radial-gradient(circle at 16% 12%, #3a4811 0 18%, transparent 36%), linear-gradient(135deg, #10130f, #242013 62%, #0c1012); font-family: Georgia, "Times New Roman", "Noto Serif SC", serif; color: var(--panel); }
    main { width: min(1080px, calc(100vw - 32px)); margin: 0 auto; padding: 42px 0 64px; }
    .hero { display: grid; grid-template-columns: 1fr; gap: 20px; }
    .title-card, .workspace-card { border: 1px solid rgba(243,234,215,.22); border-radius: 30px; padding: 34px; background: rgba(16,19,15,.76); box-shadow: 0 24px 80px rgba(0,0,0,.38); }
    .eyebrow { color: var(--accent); letter-spacing: .18em; text-transform: uppercase; font: 700 12px/1.2 Consolas, monospace; }
    h1 { margin: 20px 0 12px; font-size: clamp(42px, 7vw, 88px); line-height: .9; letter-spacing: -.06em; }
    .subtitle { max-width: 820px; color: #d9cab0; font-size: 19px; line-height: 1.7; }
    #drop-zone { margin-top: 28px; border: 3px dashed rgba(199,255,69,.72); border-radius: 34px; padding: 56px 28px; text-align: center; background: rgba(243,234,215,.09); transition: transform .15s ease, border-color .15s ease, background .15s ease; }
    #drop-zone.dragging { transform: translateY(-3px); border-color: var(--hot); background: rgba(255,106,61,.16); }
    #drop-zone strong { display: block; color: var(--accent); font-size: clamp(30px, 5vw, 58px); letter-spacing: -.04em; }
    #drop-zone span { display: block; margin-top: 12px; color: #e1d3b9; font-size: 18px; }
    .file-button { display: inline-block; margin-top: 22px; border-radius: 999px; padding: 13px 18px; background: var(--panel); color: var(--ink); cursor: pointer; font: 800 14px/1 Consolas, monospace; }
    input[type=file] { display: none; }
    .workspace-card { display: none; margin-top: 22px; color: var(--panel); }
    .workspace-card.visible { display: block; }
    .path { margin: 14px 0; padding: 14px; border-radius: 16px; background: rgba(0,0,0,.35); color: #f3ead7; font: 13px/1.5 Consolas, monospace; word-break: break-all; }
    .actions { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 18px; }
    button { border: 0; border-radius: 16px; padding: 14px 16px; cursor: pointer; background: var(--panel); color: var(--ink); font: 800 14px/1.2 Consolas, monospace; }
    button.primary { background: var(--accent); }
    button.warn { background: var(--hot); color: #fff8ed; }
    pre { margin-top: 22px; border: 1px solid rgba(243,234,215,.18); border-radius: 20px; padding: 18px; background: rgba(0,0,0,.35); color: #e8dcc8; min-height: 92px; white-space: pre-wrap; }
    @media (max-width: 760px) { .actions { grid-template-columns: 1fr; } .title-card, .workspace-card { padding: 24px; } }
  </style>
</head>
<body>
  <main>
    <section class="hero">
      <div class="title-card">
        <span class="eyebrow">DOCX XML TOOL · v2.5</span>
        <h1>拖进去，<br>自动生成工作区</h1>
        <p class="subtitle">把 .docx 拖到下面的大框里。本地 Python 服务会自动导出工作区；文档只在你电脑上处理，不上传云端。</p>
        <div id="drop-zone">
          <strong>把 .docx 拖到这里</strong>
          <span>或者点击选择 Word 文件</span>
          <label class="file-button" for="file-input">选择 .docx 文件</label>
          <input id="file-input" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
        </div>
      </div>
      <div id="workspace-card" class="workspace-card">
        <span class="eyebrow">WORKSPACE READY</span>
        <h2>工作区已生成</h2>
        <div id="workspace-path" class="path"></div>
        <div class="actions">
          <button id="open-md">打开 document.md</button>
          <button id="open-fields">打开 field_values.yaml</button>
          <button id="open-edits">打开输出目录</button>
          <button id="apply-md" class="primary">回写 Markdown</button>
          <button id="fill-fields" class="primary">填充字段</button>
          <button id="open-workspace" class="warn">打开工作区</button>
        </div>
      </div>
    </section>
    <pre id="result">等待拖入 .docx 文件...</pre>
  </main>
  <script>
    const dropZone = document.querySelector('#drop-zone');
    const fileInput = document.querySelector('#file-input');
    const result = document.querySelector('#result');
    const card = document.querySelector('#workspace-card');
    const workspacePath = document.querySelector('#workspace-path');
    let current = null;

    function show(data) {
      result.textContent = JSON.stringify(data, null, 2);
    }

    async function upload(file) {
      if (!file || !file.name.toLowerCase().endsWith('.docx')) {
        show({ ok: false, error: '请拖入 .docx 文件。' });
        return;
      }
      result.textContent = '正在导出工作区...';
      const body = new FormData();
      body.append('docx_file', file);
      const response = await fetch('/api/upload-docx', { method: 'POST', body });
      const data = await response.json();
      show(data);
      if (data.ok) {
        current = data;
        workspacePath.textContent = data.workspace;
        card.classList.add('visible');
      }
    }

    async function postForm(endpoint, values) {
      const response = await fetch(endpoint, { method: 'POST', body: new URLSearchParams(values) });
      const data = await response.json();
      show(data);
      return data;
    }

    async function openPath(path) {
      if (!path) return;
      await postForm('/api/open-path', { path });
    }

    for (const eventName of ['dragenter', 'dragover']) {
      dropZone.addEventListener(eventName, event => { event.preventDefault(); dropZone.classList.add('dragging'); });
    }
    for (const eventName of ['dragleave', 'drop']) {
      dropZone.addEventListener(eventName, event => { event.preventDefault(); dropZone.classList.remove('dragging'); });
    }
    dropZone.addEventListener('drop', event => upload(event.dataTransfer.files[0]));
    fileInput.addEventListener('change', event => upload(event.target.files[0]));

    document.querySelector('#open-md').addEventListener('click', () => openPath(current && current.document_md));
    document.querySelector('#open-fields').addEventListener('click', () => openPath(current && current.field_values));
    document.querySelector('#open-edits').addEventListener('click', () => openPath(current && current.edits_dir));
    document.querySelector('#open-workspace').addEventListener('click', () => openPath(current && current.workspace));
    document.querySelector('#apply-md').addEventListener('click', () => current && postForm('/api/apply-md', { workspace: current.workspace }));
    document.querySelector('#fill-fields').addEventListener('click', () => current && postForm('/api/fill-fields', { workspace: current.workspace }));
  </script>
</body>
</html>
"""


def run_server(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT, open_browser: bool = True, base_dir: Path | str | None = None):
    app = create_app(base_dir)

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            self._handle()

        def do_POST(self):
            self._handle()

        def log_message(self, format, *args):
            print(f"[web-gui] {self.address_string()} - {format % args}")

        def _handle(self):
            length = int(self.headers.get("Content-Length", "0") or "0")
            body = self.rfile.read(length) if length else b""
            request_headers = {key: value for key, value in self.headers.items()}
            status, headers, response_body = app.handle_request(self.command, self.path.split("?", 1)[0], body, request_headers)
            self.send_response(status)
            for key, value in headers.items():
                self.send_header(key, value)
            self.send_header("Content-Length", str(len(response_body)))
            self.end_headers()
            self.wfile.write(response_body)

    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{server.server_port}/"
    print(f"DOCX XML Tool v2.5 Web GUI: {url}")
    if open_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nWeb GUI stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    run_server()
