import json
import tempfile
import unittest
import zipfile
from pathlib import Path
import sys
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_workspace_tool import apply_markdown_workspace, export_docx_xml, fill_docx_workspace


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


def read_docx_texts(path: Path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return [node.text or "" for node in root.findall(".//w:t", NS)]


class DocxXmlExporterTest(unittest.TestCase):
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
