import json
import tempfile
import unittest
import zipfile
from pathlib import Path
import sys
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_xml_exporter import export_docx_xml, fill_docx_workspace


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
                    {
                        "label": "姓名",
                        "part": "word/document.xml",
                        "text_node_indexes": [5],
                        "old_text": "姓名：张三",
                        "new_text": "姓名：李四"
                    },
                    {
                        "label": "个人能力",
                        "part": "word/document.xml",
                        "text_node_indexes": [3],
                        "old_text": "个人能力正文",
                        "new_text": "具备较强自主学习能力"
                    }
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
            replacements.write_text(json.dumps({
                "replacements": [
                    {
                        "label": "姓名",
                        "part": "word/document.xml",
                        "text_node_indexes": [5],
                        "old_text": "姓名：王五",
                        "new_text": "姓名：李四"
                    }
                ]
            }, ensure_ascii=False), encoding="utf-8")

            with self.assertRaises(ValueError):
                fill_docx_workspace(workspace, replacements)


if __name__ == "__main__":
    unittest.main()
