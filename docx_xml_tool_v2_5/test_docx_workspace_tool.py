import json
import tempfile
import unittest
import zipfile
from pathlib import Path
import sys
from urllib.parse import urlencode
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

import docx_workspace_tool
from docx_workspace_tool import apply_markdown_workspace, export_docx_xml, fill_docx_workspace, fill_fields_workspace
from docx_markdown.fields import detect_fields_from_blocks


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def write_minimal_docx(path: Path):
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>标题</w:t></w:r></w:p>
    <w:p><w:r><w:t>自我评价正文</w:t></w:r></w:p>
    <w:p><w:r><w:t>自我评价</w:t></w:r></w:p>
    <w:p><w:r><w:t>个人能力正文</w:t></w:r></w:p>
    <w:p><w:r><w:t>个人能力</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>姓名：张三</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>第一周日志</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)


def write_complex_docx(path: Path):
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>'''
    document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>正文段落</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>横向合并</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>纵向合并</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>普通单元格</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>第二格</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>延续单元格</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>'''
    header = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>页眉文字</w:t></w:r></w:p></w:hdr>'''
    footer = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>页脚文字</w:t></w:r></w:p></w:ftr>'''
    footnotes = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>脚注文字</w:t></w:r></w:p></w:footnote></w:footnotes>'''
    endnotes = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="2"><w:p><w:r><w:t>尾注文字</w:t></w:r></w:p></w:endnote></w:endnotes>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/header1.xml", header)
        z.writestr("word/footer1.xml", footer)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/endnotes.xml", endnotes)
        z.writestr("word/media/image1.png", b"fakepngdata")


def write_table_field_docx(path: Path):
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>张三</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)


def write_content_control_docx(path: Path):
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt><w:sdtContent><w:p><w:r><w:t>请输入姓名</w:t></w:r></w:p></w:sdtContent></w:sdt>
  </w:body>
</w:document>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)


def read_docx_texts(path: Path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return [node.text or "" for node in root.findall(".//w:t", NS)]


def make_multipart_file(field_name: str, filename: str, content: bytes, content_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"):
    boundary = "----docxXmlToolBoundary"
    body = b"\r\n".join([
        f"--{boundary}".encode("utf-8"),
        f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"'.encode("utf-8"),
        f"Content-Type: {content_type}".encode("utf-8"),
        b"",
        content,
        f"--{boundary}--".encode("utf-8"),
        b"",
    ])
    return f"multipart/form-data; boundary={boundary}", body


class DocxXmlExporterTest(unittest.TestCase):
    def test_cli_starts_web_gui_with_local_options(self):
        import docx_web_gui

        calls = []
        original_run_server = docx_web_gui.run_server
        try:
            def fake_run_server(host="127.0.0.1", port=8765, open_browser=True, base_dir=None):
                calls.append({"host": host, "port": port, "open_browser": open_browser, "base_dir": base_dir})

            docx_web_gui.run_server = fake_run_server
            docx_workspace_tool.main(["--web-gui", "--host", "127.0.0.1", "--port", "0", "--no-open"])
        finally:
            docx_web_gui.run_server = original_run_server

        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0]["host"], "127.0.0.1")
        self.assertEqual(calls[0]["port"], 0)
        self.assertFalse(calls[0]["open_browser"])
        self.assertEqual(calls[0]["base_dir"], Path(__file__).resolve().parent)

    def test_gui_launcher_starts_local_web_gui(self):
        root = Path(__file__).resolve().parent
        launcher = (root / "launchers" / "04_start_web_gui.bat").read_text(encoding="utf-8")

        self.assertIn("--web-gui", launcher)
        self.assertIn("127.0.0.1", launcher)
        self.assertIn("web_gui_last_run.log", launcher)

    def test_v2_5_docs_and_package_names_are_current(self):
        root = Path(__file__).resolve().parent
        readme = (root / "README.md").read_text(encoding="utf-8")
        readme_zh = (root / "README.zh-CN.md").read_text(encoding="utf-8")
        maintenance = (root / "README_MAINTENANCE.md").read_text(encoding="utf-8")
        package_script = (root / "scripts" / "package_release.py").read_text(encoding="utf-8")

        self.assertTrue((root / "README_V2.5.md").exists())
        self.assertFalse((root / "README_V2.4.2.md").exists())
        for text in [readme, readme_zh, maintenance]:
            self.assertIn("v2.5", text)
            self.assertIn("04_start_web_gui.bat", text)
            self.assertNotIn("docx_xml_tool_v2_4_2.zip", text)
        self.assertIn("docx_xml_tool_v2_5.zip", package_script)

    def test_gui_upload_docx_exports_workspace(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "drag-source.docx"
            write_minimal_docx(source)
            content_type, body = make_multipart_file("docx_file", "drag-source.docx", source.read_bytes())
            app = create_app(tmp_path)

            status, headers, response_body = app.handle_request("POST", "/api/upload-docx", body, {"Content-Type": content_type})

            result = json.loads(response_body.decode("utf-8"))
            self.assertEqual(status, 200)
            self.assertEqual(headers["Content-Type"], "application/json; charset=utf-8")
            self.assertTrue(result["ok"])
            workspace = Path(result["workspace"])
            self.assertTrue(workspace.exists())
            self.assertTrue(Path(result["document_md"]).exists())
            self.assertTrue(Path(result["field_values"]).exists())
            self.assertEqual(Path(result["edits_dir"]), workspace / "edits")
            self.assertTrue((tmp_path / "gui_uploads" / "drag-source.docx").exists())

    def test_gui_upload_rejects_non_docx_file(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            content_type, body = make_multipart_file("docx_file", "notes.txt", b"not a docx", "text/plain")
            app = create_app(tmp_path)

            status, headers, response_body = app.handle_request("POST", "/api/upload-docx", body, {"Content-Type": content_type})

            result = json.loads(response_body.decode("utf-8"))
            self.assertEqual(status, 400)
            self.assertFalse(result["ok"])
            self.assertIn(".docx", result["error"])

    def test_gui_home_page_has_drag_drop_workflow(self):
        from docx_web_gui import create_app

        app = create_app(Path.cwd())
        status, headers, html_body = app.handle_request("GET", "/", b"")
        html = html_body.decode("utf-8")

        self.assertEqual(status, 200)
        self.assertIn("把 .docx 拖到这里", html)
        self.assertIn("drop-zone", html)
        self.assertIn("打开 document.md", html)
        self.assertIn("回写 Markdown", html)
        self.assertIn("打开输出目录", html)

    def test_gui_open_path_allows_base_dir_path_without_launching_in_tests(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            target = tmp_path / "document.md"
            target.write_text("hello", encoding="utf-8")
            app = create_app(tmp_path)
            payload = urlencode({"path": str(target), "test_mode": "1"}).encode("utf-8")

            status, headers, response_body = app.handle_request("POST", "/api/open-path", payload)

            result = json.loads(response_body.decode("utf-8"))
            self.assertEqual(status, 200)
            self.assertTrue(result["ok"])
            self.assertEqual(Path(result["path"]), target)

    def test_gui_open_path_rejects_outside_base_dir(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            outside = tmp_path.parent / "outside.md"
            app = create_app(tmp_path)
            payload = urlencode({"path": str(outside), "test_mode": "1"}).encode("utf-8")

            status, headers, response_body = app.handle_request("POST", "/api/open-path", payload)

            result = json.loads(response_body.decode("utf-8"))
            self.assertEqual(status, 400)
            self.assertFalse(result["ok"])
            self.assertIn("工作区", result["error"])

    def test_gui_home_page_and_health_endpoint(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            app = create_app(Path(tmp))
            html_status, html_headers, html_body = app.handle_request("GET", "/", b"")
            health_status, health_headers, health_body = app.handle_request("GET", "/api/health", b"")

        self.assertEqual(html_status, 200)
        self.assertEqual(html_headers["Content-Type"], "text/html; charset=utf-8")
        self.assertIn("DOCX XML Tool v2.5", html_body.decode("utf-8"))
        self.assertIn("导出工作区", html_body.decode("utf-8"))
        self.assertEqual(health_status, 200)
        self.assertEqual(health_headers["Content-Type"], "application/json; charset=utf-8")
        health = json.loads(health_body.decode("utf-8"))
        self.assertEqual(health["status"], "ok")
        self.assertEqual(health["version"], "2.5")

    def test_gui_export_action_creates_workspace(self):
        from docx_web_gui import create_app

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "gui-source.docx"
            workspace_root = tmp_path / "workspaces"
            write_minimal_docx(source)
            app = create_app(tmp_path)
            payload = urlencode({"input_docx": str(source), "workspace_root": str(workspace_root)}).encode("utf-8")

            status, headers, body = app.handle_request("POST", "/api/export", payload)

            result = json.loads(body.decode("utf-8"))
            self.assertEqual(status, 200)
            self.assertEqual(headers["Content-Type"], "application/json; charset=utf-8")
            self.assertTrue(result["ok"])
            workspace = Path(result["workspace"])
            self.assertTrue((workspace / "document.md").exists())
            self.assertTrue((workspace / "content_map.json").exists())
            self.assertTrue((workspace / "fields.yaml").exists())
            self.assertIn("gui-source", workspace.name)

    def test_field_fill_launcher_and_readme_explain_workspace_input(self):
        root = Path(__file__).resolve().parent
        launcher = (root / "launchers" / "03_fill_fields_to_docx.bat").read_text(encoding="utf-8")
        readme_zh = (root / "README.zh-CN.md").read_text(encoding="utf-8")

        self.assertIn("工作区文件夹", launcher)
        self.assertIn("不能直接拖原始 .docx", launcher)
        self.assertIn("先运行 01_export_docx_workspace.bat", launcher)
        self.assertIn("03_fill_fields_to_docx.bat", readme_zh)
        self.assertIn("不能直接拖原始 .docx", readme_zh)
        self.assertIn("把整个工作区文件夹拖到", readme_zh)

    def test_export_generates_fields_and_field_values_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            out = tmp_path / "exported"
            write_minimal_docx(source)

            export_docx_xml(source, out)

            fields = (out / "fields.yaml").read_text(encoding="utf-8")
            values = (out / "field_values.yaml").read_text(encoding="utf-8")
            self.assertIn("key: name", fields)
            self.assertIn("label: 姓名", fields)
            self.assertIn("current_text: 张三", fields)
            self.assertIn("name:", values)
            self.assertIn("# 姓名", values)

    def test_detect_fields_from_label_value_blocks(self):
        blocks = [
            {"block_id": "b00001", "kind": "paragraph", "part": "word/document.xml", "path": "p[0]", "text": "姓名：张三"},
            {"block_id": "b00002", "kind": "paragraph", "part": "word/document.xml", "path": "p[1]", "text": "普通正文"},
        ]

        fields = detect_fields_from_blocks(blocks)

        self.assertEqual(len(fields), 1)
        self.assertEqual(fields[0]["key"], "name")
        self.assertEqual(fields[0]["label"], "姓名")
        self.assertEqual(fields[0]["current_text"], "张三")
        self.assertEqual(fields[0]["replacement_mode"], "replace_value")
        self.assertEqual(fields[0]["locations"][0]["block_id"], "b00001")

    def test_field_values_plan_replaces_only_label_value_part(self):
        from docx_markdown.fields import parse_field_values, plan_field_replacements

        fields = [{
            "key": "name",
            "label": "姓名",
            "current_text": "张三",
            "replacement_mode": "replace_value",
            "risk": "low",
            "locations": [{"block_id": "b00001", "part": "word/document.xml", "kind": "paragraph", "path": "p[0]"}],
        }]
        blocks = [{"block_id": "b00001", "text": "姓名：张三"}]
        values = parse_field_values("name: 李四\n")

        plan = plan_field_replacements(fields, values, blocks)

        self.assertEqual(plan["replacements"], {"b00001": "姓名：李四"})
        self.assertEqual(plan["changed_fields"][0]["key"], "name")
        self.assertEqual(plan["warnings"], [])

    def test_field_fill_skips_ambiguous_duplicate_locations(self):
        from docx_markdown.fields import plan_field_replacements

        fields = [{
            "key": "name",
            "label": "姓名",
            "current_text": "张三",
            "replacement_mode": "replace_value",
            "risk": "low",
            "locations": [
                {"block_id": "b00001", "part": "word/document.xml", "kind": "paragraph", "path": "p[0]"},
                {"block_id": "b00002", "part": "word/document.xml", "kind": "paragraph", "path": "p[1]"},
            ],
        }]
        blocks = [
            {"block_id": "b00001", "text": "姓名：张三"},
            {"block_id": "b00002", "text": "姓名：张三"},
        ]

        plan = plan_field_replacements(fields, {"name": "李四"}, blocks)

        self.assertEqual(plan["replacements"], {})
        self.assertEqual(len(plan["warnings"]), 1)
        self.assertIn("ambiguous", plan["warnings"][0])

    def test_export_detects_table_adjacent_label_value_field(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "table-field.docx"
            workspace = tmp_path / "exported"
            write_table_field_docx(source)

            export_docx_xml(source, workspace)

            fields = (workspace / "fields.yaml").read_text(encoding="utf-8")
            self.assertIn("key: name", fields)
            self.assertIn("label: 姓名", fields)
            self.assertIn("current_text: 张三", fields)
            self.assertIn("replacement_mode: replace_all", fields)

    def test_content_control_placeholder_is_exported_as_high_risk_field(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "content-control.docx"
            workspace = tmp_path / "exported"
            write_content_control_docx(source)

            export_docx_xml(source, workspace)

            content_map = json.loads((workspace / "content_map.json").read_text(encoding="utf-8"))
            self.assertTrue(any(block["kind"] == "content_control" for block in content_map["blocks"]))
            fields = (workspace / "fields.yaml").read_text(encoding="utf-8")
            self.assertIn("risk: high", fields)
            self.assertIn("请输入姓名", fields)

    def test_exports_docx_xml_structure_and_text_index(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            out = tmp_path / "exported"
            write_minimal_docx(source)

            result = export_docx_xml(source, out)

            self.assertEqual(result, out)
            self.assertTrue((out / "original.docx").exists())
            self.assertTrue((out / "unpacked" / "word" / "document.xml").exists())
            self.assertTrue((out / "pretty_xml" / "word" / "document.xml").exists())
            self.assertIn("标题", (out / "structure.md").read_text(encoding="utf-8"))
            text_index = json.loads((out / "text_index.json").read_text(encoding="utf-8"))
            self.assertTrue(any(item["text"] == "姓名：张三" and item["kind"] == "table_cell" for item in text_index))
            manifest = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
            self.assertTrue(any(item["path"] == "word/document.xml" for item in manifest))
            fields_text = (out / "candidate_fields.yaml").read_text(encoding="utf-8")
            self.assertIn("姓名", fields_text)

    def test_writes_section_context_for_heading_neighbors(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            out = tmp_path / "exported"
            write_minimal_docx(source)

            export_docx_xml(source, out)

            context = (out / "section_context.md").read_text(encoding="utf-8")
            self.assertIn("## 个人能力", context)
            self.assertIn("个人能力正文", context)
            self.assertIn("自我评价正文", context)
            self.assertIn("Likely content direction: before_heading", context)
            self.assertIn("Warning", context)

    def test_default_output_uses_timestamped_workspace_and_index(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            workspace = tmp_path / "outputs"
            write_minimal_docx(source)

            result = export_docx_xml(source, workspace_root=workspace)

            self.assertEqual(result.parent, workspace)
            self.assertIn("示例", result.name)
            self.assertTrue((result / "structure.md").exists())
            index = (workspace / "index.md").read_text(encoding="utf-8")
            self.assertIn("示例.docx", index)
            self.assertIn(result.name, index)

    def test_fill_fields_workspace_writes_docx_and_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            workspace = tmp_path / "exported"
            write_minimal_docx(source)
            export_docx_xml(source, workspace)
            (workspace / "field_values.yaml").write_text("name: 李四\n", encoding="utf-8")

            output = fill_fields_workspace(workspace)

            self.assertTrue(output.exists())
            texts = read_docx_texts(output)
            self.assertIn("姓名：李四", texts)
            self.assertNotIn("姓名：张三", texts)
            audit = (workspace / "edits" / "field_fill_audit_001.md").read_text(encoding="utf-8")
            self.assertIn("changed_fields: 1", audit)
            self.assertIn("name", audit)
            self.assertIn("姓名：张三", audit)
            self.assertIn("姓名：李四", audit)
            self.assertIn("zip_test: ok", audit)

    def test_fill_replaces_explicit_text_nodes_and_writes_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            workspace = tmp_path / "exported"
            replacements = tmp_path / "replacements.json"
            write_minimal_docx(source)
            export_docx_xml(source, workspace)
            replacements.write_text(json.dumps({
                "replacements": [
                    {"label": "姓名", "part": "word/document.xml", "text_node_indexes": [5], "old_text": "姓名：张三", "new_text": "姓名：李四"},
                    {"label": "个人能力", "part": "word/document.xml", "text_node_indexes": [3], "old_text": "个人能力正文", "new_text": "具备较强自主学习能力"}
                ]
            }, ensure_ascii=False), encoding="utf-8")

            output_docx = fill_docx_workspace(workspace, replacements)

            self.assertTrue(output_docx.exists())
            texts = read_docx_texts(output_docx)
            self.assertIn("姓名：李四", texts)
            self.assertIn("具备较强自主学习能力", texts)
            self.assertIn("自我评价正文", texts)
            self.assertNotIn("姓名：张三", texts)
            audit = sorted((workspace / "edits").glob("edit_audit_*.md"))[0].read_text(encoding="utf-8")
            self.assertIn("label: 姓名", audit)
            self.assertIn("old_text matched: True", audit)
            self.assertIn("zip_test: ok", audit)

    def test_fill_refuses_when_old_text_does_not_match(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "示例.docx"
            workspace = tmp_path / "exported"
            replacements = tmp_path / "bad_replacements.json"
            write_minimal_docx(source)
            export_docx_xml(source, workspace)
            replacements.write_text(json.dumps({"replacements": [{"label": "姓名", "part": "word/document.xml", "text_node_indexes": [5], "old_text": "姓名：王五", "new_text": "姓名：李四"}]}, ensure_ascii=False), encoding="utf-8")

            with self.assertRaises(ValueError):
                fill_docx_workspace(workspace, replacements)

    def test_markdown_export_and_apply_preserves_old_workspace_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "sample.docx"
            workspace = tmp_path / "exported"
            write_minimal_docx(source)
            export_docx_xml(source, workspace)

            self.assertTrue((workspace / "unpacked" / "word" / "document.xml").exists())
            self.assertTrue((workspace / "pretty_xml" / "word" / "document.xml").exists())
            self.assertTrue((workspace / "document.md").exists())
            self.assertTrue((workspace / "content_map.json").exists())
            self.assertTrue((workspace / "format_profile.yaml").exists())

            text_index = json.loads((workspace / "text_index.json").read_text(encoding="utf-8"))
            old_text = next(item["text"] for item in text_index if item["kind"] == "table_cell")
            new_text = old_text + " updated"
            md = workspace / "document.md"
            md.write_text(md.read_text(encoding="utf-8").replace(old_text, new_text), encoding="utf-8")

            output_docx = apply_markdown_workspace(workspace)

            self.assertTrue(output_docx.exists())
            self.assertTrue((workspace / "edits" / "markdown_apply_audit_001.md").exists())
            texts = read_docx_texts(output_docx)
            self.assertIn(new_text, texts)
            self.assertNotIn(old_text, texts)

    def test_markdown_apply_rejects_missing_required_block_marker(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            text = md.read_text(encoding="utf-8")
            marker_line = next(line for line in text.splitlines() if line.startswith("<!--docx:block"))
            md.write_text(text.replace(marker_line + "\n", "", 1), encoding="utf-8")

            with self.assertRaisesRegex(Exception, "缺少必要"):
                apply_markdown_workspace(workspace)

    def test_markdown_apply_rejects_duplicate_block_marker(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            text = md.read_text(encoding="utf-8")
            marker_line = next(line for line in text.splitlines() if line.startswith("<!--docx:block"))
            md.write_text(text + "\n" + marker_line + "\n重复内容\n", encoding="utf-8")

            with self.assertRaisesRegex(Exception, "重复"):
                apply_markdown_workspace(workspace)

    def test_markdown_apply_rejects_unknown_block_marker(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            md.write_text(md.read_text(encoding="utf-8") + "\n<!--docx:block b99999-->\n未知内容\n", encoding="utf-8")

            with self.assertRaisesRegex(Exception, "无法识别"):
                apply_markdown_workspace(workspace)

    def test_markdown_export_and_apply_headers_footers_footnotes_endnotes(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            text = md.read_text(encoding="utf-8")
            self.assertIn("页眉文字", text)
            self.assertIn("页脚文字", text)
            self.assertIn("脚注文字", text)
            self.assertIn("尾注文字", text)
            md.write_text(text.replace("页眉文字", "新页眉").replace("页脚文字", "新页脚").replace("脚注文字", "新脚注").replace("尾注文字", "新尾注"), encoding="utf-8")

            output = apply_markdown_workspace(workspace)

            with zipfile.ZipFile(output) as z:
                self.assertIn("新页眉", z.read("word/header1.xml").decode("utf-8"))
                self.assertIn("新页脚", z.read("word/footer1.xml").decode("utf-8"))
                self.assertIn("新脚注", z.read("word/footnotes.xml").decode("utf-8"))
                self.assertIn("新尾注", z.read("word/endnotes.xml").decode("utf-8"))

    def test_table_merge_metadata_is_recorded(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)

            content_map = json.loads((workspace / "content_map.json").read_text(encoding="utf-8"))
            merges = content_map["table_merges"][0]["merges"]
            self.assertTrue(any(item["gridSpan"] == 2 for item in merges))
            self.assertTrue(any(item["vMerge"] == "restart" for item in merges))
            self.assertTrue(any(item["vMerge"] == "continue" for item in merges))

    def test_markdown_apply_preserves_media_parts(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            md.write_text(md.read_text(encoding="utf-8").replace("正文段落", "正文修改"), encoding="utf-8")

            output = apply_markdown_workspace(workspace)

            with zipfile.ZipFile(output) as z:
                self.assertEqual(z.read("word/media/image1.png"), b"fakepngdata")

    def test_markdown_apply_audit_lists_changed_blocks(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex.docx"
            workspace = tmp_path / "exported"
            write_complex_docx(source)
            export_docx_xml(source, workspace)
            md = workspace / "document.md"
            md.write_text(md.read_text(encoding="utf-8").replace("正文段落", "正文修改"), encoding="utf-8")

            apply_markdown_workspace(workspace)

            audit = (workspace / "edits" / "markdown_apply_audit_001.md").read_text(encoding="utf-8")
            self.assertIn("changed_blocks: 1", audit)
            self.assertIn("## Changed blocks", audit)
            self.assertIn("正文段落", audit)
            self.assertIn("正文修改", audit)


if __name__ == "__main__":
    unittest.main()
