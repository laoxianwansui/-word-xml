from __future__ import annotations

import html
import json
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {
    "w": W_NS,
    "r": R_NS,
    "wp": WP_NS,
    "a": A_NS,
    "pic": PIC_NS,
    "rel": REL_NS,
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def xml_escape(value: str) -> str:
    """Escape XML special characters consistently."""
    return html.escape(value or "", quote=True).replace("'", "&apos;")


def xml_unescape(value: str) -> str:
    """Reverse XML entity escaping, including apostrophes."""
    return html.unescape(value or "")


class DocxCoreError(Exception):
    """User-facing DOCX core error with a Chinese message."""


@dataclass
class Block:
    block_id: str
    kind: str
    part: str
    text: str
    paragraph_index: Optional[int] = None
    table_index: Optional[int] = None
    row_index: Optional[int] = None
    col_index: Optional[int] = None
    path: str = ""
    metadata: Dict[str, object] = field(default_factory=dict)


@dataclass
class ExportResult:
    markdown: str
    content_map: Dict[str, object]
    format_profile: Dict[str, object]


class DocxCore:
    def export(self, input_docx: Path, output_dir: Path) -> ExportResult:
        input_docx = Path(input_docx)
        output_dir = Path(output_dir)
        self._validate_docx(input_docx)

        try:
            with zipfile.ZipFile(input_docx) as archive:
                document_xml = self._read_required_xml(archive, "word/document.xml")
                parts = self._discover_parts(archive)
                parsed_parts = [("word/document.xml", document_xml)]
                parsed_parts.extend((part, self._read_optional_xml(archive, part)) for part in parts)
        except zipfile.BadZipFile as exc:
            raise DocxCoreError("文件不是有效的DOCX压缩包，可能已损坏。") from exc

        blocks: List[Block] = []
        table_profiles: List[Dict[str, object]] = []
        md_sections: List[str] = []

        for part_name, root in parsed_parts:
            if root is None:
                continue
            section_title = self._part_title(part_name)
            part_blocks, part_tables_md, part_tables_profile = self._parse_part(root, part_name, len(blocks))
            blocks.extend(part_blocks)
            table_profiles.extend(part_tables_profile)
            rendered = self._render_part_markdown(section_title, part_blocks, part_tables_md)
            if rendered.strip():
                md_sections.append(rendered)

        markdown = "\n\n".join(md_sections).strip() + "\n"
        content_map = {
            "version": "2.4.0",
            "source": str(input_docx),
            "blocks": [block.__dict__ for block in blocks],
            "table_merges": table_profiles,
            "notes": [
                "Markdown中的 <!--docx:block ...--> 注释用于稳定回写定位。",
                "DOCX XML由结构化解析器读写，特殊字符由XML序列化自动转义。",
            ],
        }
        format_profile = {
            "version": "2.4.0",
            "style_policy": "semantic_markdown_only",
            "preserved_by_apply": [
                "docx package relationships",
                "paragraph and run style containers",
                "tables including gridSpan/vMerge metadata",
                "headers",
                "footers",
                "footnotes",
            ],
            "stripped_from_markdown": [
                "Word direct formatting noise",
                "run-level visual styles",
                "theme/style ids not needed for semantic editing",
            ],
            "table_policy": "render aligned markdown grid; store merge metadata in content_map.json",
        }

        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / "document.md").write_text(markdown, encoding="utf-8")
        (output_dir / "content_map.json").write_text(
            json.dumps(content_map, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (output_dir / "format_profile.yaml").write_text(self._to_simple_yaml(format_profile), encoding="utf-8")
        return ExportResult(markdown, content_map, format_profile)

    def apply(self, original_docx: Path, edited_markdown: Path, content_map: Path, output_docx: Path) -> None:
        original_docx = Path(original_docx)
        edited_markdown = Path(edited_markdown)
        content_map = Path(content_map)
        output_docx = Path(output_docx)
        self._validate_docx(original_docx)
        if not edited_markdown.exists() or edited_markdown.stat().st_size == 0:
            raise DocxCoreError("Markdown文件为空或不存在，无法回写。")
        if not content_map.exists() or content_map.stat().st_size == 0:
            raise DocxCoreError("content_map.json为空或不存在，无法定位回写内容。")

        try:
            mapping = json.loads(content_map.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise DocxCoreError("content_map.json格式损坏，无法读取回写映射。") from exc

        replacements = self._extract_markdown_blocks(edited_markdown.read_text(encoding="utf-8"))
        block_defs = {item["block_id"]: item for item in mapping.get("blocks", [])}
        touched_parts = {item["part"] for item in block_defs.values() if item["block_id"] in replacements}
        if not touched_parts:
            raise DocxCoreError("Markdown中没有找到可回写的docx:block标记。")

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        try:
            replacements_by_part: Dict[str, bytes] = {}
            with zipfile.ZipFile(original_docx, mode="r") as archive:
                for part in sorted(touched_parts):
                    root = self._read_required_xml(archive, part)
                    self._apply_part(root, part, block_defs, replacements)
                    replacements_by_part[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                self._rewrite_zip(archive, output_docx, replacements_by_part)
        except zipfile.BadZipFile as exc:
            raise DocxCoreError("输出DOCX写入失败，原文件可能已损坏。") from exc

    def _validate_docx(self, path: Path) -> None:
        if not path.exists():
            raise DocxCoreError(f"文件不存在：{path}")
        if path.stat().st_size == 0:
            raise DocxCoreError("DOCX文件为空，无法解析。")
        if path.suffix.lower() != ".docx":
            raise DocxCoreError("输入文件不是.docx格式。")
        try:
            with zipfile.ZipFile(path) as archive:
                if "word/document.xml" not in archive.namelist():
                    raise DocxCoreError("DOCX缺失核心文件 word/document.xml，无法解析。")
        except zipfile.BadZipFile as exc:
            raise DocxCoreError("文件不是有效的DOCX压缩包，可能已损坏。") from exc

    def _read_required_xml(self, archive: zipfile.ZipFile, name: str) -> ET.Element:
        try:
            raw = archive.read(name)
        except KeyError as exc:
            raise DocxCoreError(f"DOCX缺失核心XML：{name}") from exc
        try:
            return ET.fromstring(raw)
        except ET.ParseError as exc:
            raise DocxCoreError(f"XML解析失败：{name}，文件可能已损坏。") from exc

    def _read_optional_xml(self, archive: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
        try:
            raw = archive.read(name)
        except KeyError:
            return None
        try:
            return ET.fromstring(raw)
        except ET.ParseError as exc:
            raise DocxCoreError(f"XML解析失败：{name}，文件可能已损坏。") from exc

    def _discover_parts(self, archive: zipfile.ZipFile) -> List[str]:
        names = set(archive.namelist())
        parts = [
            name for name in names
            if re.match(r"word/(header|footer)\d+\.xml$", name) or name == "word/footnotes.xml"
        ]
        return sorted(parts)

    def _part_title(self, part_name: str) -> str:
        if part_name == "word/document.xml":
            return "正文"
        if "/header" in part_name:
            return "页眉"
        if "/footer" in part_name:
            return "页脚"
        if part_name.endswith("footnotes.xml"):
            return "脚注"
        return part_name

    def _parse_part(self, root: ET.Element, part_name: str, block_offset: int) -> Tuple[List[Block], List[Tuple[str, str]], List[Dict[str, object]]]:
        blocks: List[Block] = []
        table_markdown: List[Tuple[str, str]] = []
        table_profiles: List[Dict[str, object]] = []
        body = root.find("w:body", NS)
        body_like = body if body is not None else root
        block_counter = block_offset
        paragraph_index = 0
        table_index = 0

        for child in list(body_like):
            if child.tag == w_tag("p"):
                text = self._element_text(child)
                if text.strip():
                    block_id = f"b{block_counter:05d}"
                    blocks.append(Block(block_id, "paragraph", part_name, text, paragraph_index=paragraph_index, path=f"p[{paragraph_index}]"))
                    block_counter += 1
                paragraph_index += 1
            elif child.tag == w_tag("tbl"):
                grid, cell_blocks, profile = self._parse_table(child, part_name, table_index, block_counter)
                block_counter += len(cell_blocks)
                blocks.extend(cell_blocks)
                table_markdown.append((f"table-{table_index}", self._render_table(grid)))
                table_profiles.append(profile)
                table_index += 1

        if part_name.endswith("footnotes.xml"):
            footnote_blocks, next_counter = self._parse_footnotes(root, part_name, block_counter)
            blocks.extend(footnote_blocks)
            block_counter = next_counter
        return blocks, table_markdown, table_profiles

    def _parse_footnotes(self, root: ET.Element, part_name: str, block_counter: int) -> Tuple[List[Block], int]:
        blocks: List[Block] = []
        for note in root.findall("w:footnote", NS):
            note_id = note.attrib.get(w_tag("id"), "")
            if note_id in {"-1", "0"}:
                continue
            text = self._element_text(note)
            if text.strip():
                block_id = f"b{block_counter:05d}"
                blocks.append(Block(block_id, "footnote", part_name, text, path=f"footnote[{note_id}]", metadata={"footnote_id": note_id}))
                block_counter += 1
        return blocks, block_counter

    def _parse_table(self, tbl: ET.Element, part_name: str, table_index: int, block_counter: int) -> Tuple[List[List[Optional[str]]], List[Block], Dict[str, object]]:
        rows = tbl.findall("w:tr", NS)
        grid: List[List[Optional[str]]] = []
        blocks: List[Block] = []
        merges: List[Dict[str, object]] = []
        active_vmerge: Dict[int, str] = {}

        for row_idx, row in enumerate(rows):
            grid_row: List[Optional[str]] = []
            col_idx = 0
            for tc in row.findall("w:tc", NS):
                tc_pr = tc.find("w:tcPr", NS)
                grid_span = 1
                vmerge_value = None
                if tc_pr is not None:
                    span = tc_pr.find("w:gridSpan", NS)
                    if span is not None:
                        try:
                            grid_span = max(1, int(span.attrib.get(w_tag("val"), "1")))
                        except ValueError:
                            grid_span = 1
                    vmerge = tc_pr.find("w:vMerge", NS)
                    if vmerge is not None:
                        vmerge_value = vmerge.attrib.get(w_tag("val"), "continue")

                text = self._element_text(tc)
                if vmerge_value == "continue":
                    for span_col in range(col_idx, col_idx + grid_span):
                        active_vmerge[span_col] = active_vmerge.get(span_col, "")
                    cell_text = ""
                else:
                    cell_text = text
                    if vmerge_value == "restart":
                        for span_col in range(col_idx, col_idx + grid_span):
                            active_vmerge[span_col] = text
                    else:
                        for span_col in range(col_idx, col_idx + grid_span):
                            active_vmerge.pop(span_col, None)

                block_id = f"b{block_counter:05d}"
                blocks.append(Block(
                    block_id,
                    "table_cell",
                    part_name,
                    cell_text,
                    table_index=table_index,
                    row_index=row_idx,
                    col_index=col_idx,
                    path=f"tbl[{table_index}]/tr[{row_idx}]/tc[{col_idx}]",
                    metadata={"grid_span": grid_span, "v_merge": vmerge_value},
                ))
                block_counter += 1

                grid_row.append(cell_text)
                for _ in range(grid_span - 1):
                    grid_row.append("")
                if grid_span > 1 or vmerge_value:
                    merges.append({"row": row_idx, "col": col_idx, "gridSpan": grid_span, "vMerge": vmerge_value})
                col_idx += grid_span

            max_cols = max(len(grid_row), len(grid[0]) if grid else 0)
            while len(grid_row) < max_cols:
                grid_row.append("")
            for old_row in grid:
                while len(old_row) < max_cols:
                    old_row.append("")
            grid.append(grid_row)

        return grid, blocks, {"part": part_name, "table_index": table_index, "merges": merges, "rows": len(grid), "cols": max((len(r) for r in grid), default=0)}

    def _render_part_markdown(self, title: str, blocks: List[Block], tables: List[Tuple[str, str]]) -> str:
        lines = [f"## {title}"]
        table_by_index = {int(name.split("-")[1]): md for name, md in tables}
        emitted_tables = set()
        for block in blocks:
            if block.kind == "table_cell":
                table_index = int(block.table_index or 0)
                if table_index not in emitted_tables:
                    lines.append(f"\n<!--docx:table {table_index}-->\n{table_by_index.get(table_index, '')}".rstrip())
                    emitted_tables.add(table_index)
                lines.append(f"<!--docx:block {block.block_id}-->")
                lines.append(block.text.strip())
                continue
            lines.append(f"\n<!--docx:block {block.block_id}-->")
            prefix = "> " if block.kind == "footnote" else ""
            lines.append(prefix + block.text.strip())
        return "\n".join(lines).strip()

    def _render_table(self, grid: List[List[Optional[str]]]) -> str:
        if not grid:
            return ""
        col_count = max(len(row) for row in grid)
        normalized = [[self._clean_table_cell(row[i] if i < len(row) else "") for i in range(col_count)] for row in grid]
        lines = ["| " + " | ".join(normalized[0]) + " |"]
        lines.append("| " + " | ".join(["---"] * col_count) + " |")
        for row in normalized[1:]:
            lines.append("| " + " | ".join(row) + " |")
        return "\n".join(lines)

    def _clean_table_cell(self, value: Optional[str]) -> str:
        return (value or "").replace("\n", "<br>").replace("|", "\\|").strip()

    def _element_text(self, element: ET.Element) -> str:
        fragments: List[str] = []
        for node in element.iter():
            if node.tag == w_tag("t"):
                fragments.append(node.text or "")
            elif node.tag == w_tag("tab"):
                fragments.append("\t")
            elif node.tag == w_tag("br"):
                fragments.append("\n")
        return "".join(fragments)

    def _extract_markdown_blocks(self, markdown: str) -> Dict[str, str]:
        marker = re.compile(r"^<!--docx:block\s+([A-Za-z0-9_-]+)-->", re.MULTILINE)
        matches = list(marker.finditer(markdown))
        replacements: Dict[str, str] = {}
        for idx, match in enumerate(matches):
            block_id = match.group(1)
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(markdown)
            raw = markdown[start:end]
            raw = re.sub(r"\n?<!--docx:table\s+\d+-->.*?(?=\n<!--docx:block|\Z)", "", raw, flags=re.S)
            text = raw.strip()
            if text.startswith("> "):
                text = "\n".join(line[2:] if line.startswith("> ") else line for line in text.splitlines())
            replacements[block_id] = text.replace("<br>", "\n")
        return replacements

    def _apply_part(self, root: ET.Element, part: str, block_defs: Dict[str, Dict[str, object]], replacements: Dict[str, str]) -> None:
        body = root.find("w:body", NS)
        body_like = body if body is not None else root
        paragraphs = [child for child in list(body_like) if child.tag == w_tag("p")]
        tables = [child for child in list(body_like) if child.tag == w_tag("tbl")]

        for block_id, text in replacements.items():
            block = block_defs.get(block_id)
            if not block or block.get("part") != part:
                continue
            kind = block.get("kind")
            if kind == "paragraph":
                index = block.get("paragraph_index")
                if isinstance(index, int) and index < len(paragraphs):
                    self._replace_element_text(paragraphs[index], text)
            elif kind == "table_cell":
                table_index = block.get("table_index")
                row_index = block.get("row_index")
                col_index = block.get("col_index")
                tc = self._find_table_cell(tables, table_index, row_index, col_index)
                if tc is not None:
                    self._replace_element_text(tc, text)
            elif kind == "footnote":
                footnote_id = str(block.get("metadata", {}).get("footnote_id", ""))
                note = self._find_footnote(root, footnote_id)
                if note is not None:
                    self._replace_element_text(note, text)

    def _find_table_cell(self, tables: List[ET.Element], table_index: object, row_index: object, col_index: object) -> Optional[ET.Element]:
        if not all(isinstance(v, int) for v in [table_index, row_index, col_index]):
            return None
        if table_index >= len(tables):
            return None
        rows = tables[table_index].findall("w:tr", NS)
        if row_index >= len(rows):
            return None
        cells = rows[row_index].findall("w:tc", NS)
        logical_col = 0
        for cell in cells:
            if logical_col == col_index:
                return cell
            span = cell.find("w:tcPr/w:gridSpan", NS)
            try:
                logical_col += int(span.attrib.get(w_tag("val"), "1")) if span is not None else 1
            except ValueError:
                logical_col += 1
        return None

    def _rewrite_zip(self, source: zipfile.ZipFile, output_docx: Path, replacements: Dict[str, bytes]) -> None:
        with zipfile.ZipFile(output_docx, mode="w") as target:
            for info in source.infolist():
                data = replacements.get(info.filename)
                if data is None:
                    data = source.read(info.filename)
                target.writestr(info, data)

    def _find_footnote(self, root: ET.Element, footnote_id: str) -> Optional[ET.Element]:
        for note in root.findall("w:footnote", NS):
            if note.attrib.get(w_tag("id"), "") == footnote_id:
                return note
        return None

    def _replace_element_text(self, element: ET.Element, new_text: str) -> None:
        text_nodes = [node for node in element.iter() if node.tag == w_tag("t")]
        if not text_nodes:
            paragraph = element.find(".//w:p", NS) if element.tag != w_tag("p") else element
            if paragraph is None:
                paragraph = ET.SubElement(element, w_tag("p"))
            run = ET.SubElement(paragraph, w_tag("r"))
            text_node = ET.SubElement(run, w_tag("t"))
            text_nodes = [text_node]
        text_nodes[0].text = new_text
        if new_text.startswith(" ") or new_text.endswith(" "):
            text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for node in text_nodes[1:]:
            node.text = ""

    def _to_simple_yaml(self, value: Dict[str, object]) -> str:
        lines: List[str] = []
        for key, item in value.items():
            if isinstance(item, list):
                lines.append(f"{key}:")
                for entry in item:
                    lines.append(f"  - {entry}")
            else:
                lines.append(f"{key}: {item}")
        return "\n".join(lines) + "\n"
