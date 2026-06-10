#!/usr/bin/env python3
import argparse
import json
import re
import shutil
import sys
import warnings
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W_NS}
XML_EXTENSIONS = {".xml", ".rels"}
MEDIA_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".mp4", ".mp3", ".wav"}
CHINESE_RE = re.compile(r"[一-鿿]")
LABEL_VALUE_RE = re.compile(r"^\s*([一-鿿A-Za-z0-9_（）()]{1,12})\s*[:：]\s*(.+?)\s*$")
PLACEHOLDER_RE = re.compile(r"(XXX+|_{2,}|—{2,}|请填写|请输入|待填写)")
MARKER_RE = re.compile(r"^[<《》>\-—_\s]+$")
SECTION_TITLES = {
    "个人信息", "教育背景", "主修课程", "个人能力", "自我评价", "获奖情况", "实习经验",
    "工作经历", "项目经历", "校园经历", "技能证书", "联系方式", "基本信息", "求职意向",
}


def contains_chinese(text):
    return bool(CHINESE_RE.search(text or ""))


def qn(local):
    return f"{{{W_NS}}}{local}"


def validate_docx(path):
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在：{path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"仅支持 .docx 文件：{path}")
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise ValueError("不是标准 DOCX 文件：缺少核心 XML")
            bad = z.testzip()
            if bad:
                raise ValueError(f"DOCX 压缩包损坏：{bad}")
    except zipfile.BadZipFile as exc:
        raise ValueError("文件不是有效 ZIP/DOCX") from exc


def read_content_types(unpacked_dir):
    content_types = {}
    path = unpacked_dir / "[Content_Types].xml"
    if not path.exists():
        return content_types
    try:
        root = ET.parse(path).getroot()
        for override in root.findall(f"{{{PKG_CT_NS}}}Override"):
            part_name = (override.get("PartName") or "").lstrip("/")
            if part_name:
                content_types[part_name] = override.get("ContentType", "")
        defaults = {d.get("Extension", "").lower(): d.get("ContentType", "") for d in root.findall(f"{{{PKG_CT_NS}}}Default")}
        content_types["__defaults__"] = defaults
    except Exception as exc:
        warnings.warn(f"读取 [Content_Types].xml 失败：{exc}")
    return content_types


def file_type(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".rels":
        return "rels"
    if suffix == ".xml":
        return "xml"
    if suffix in MEDIA_EXTENSIONS:
        return "media"
    return "other"


def content_type_for(rel_path, content_types):
    if rel_path in content_types:
        return content_types[rel_path]
    suffix = Path(rel_path).suffix.lower().lstrip(".")
    return content_types.get("__defaults__", {}).get(suffix, "")


def write_pretty_xml(src, dst):
    try:
        tree = ET.parse(src)
        ET.indent(tree, space="  ")
        dst.parent.mkdir(parents=True, exist_ok=True)
        tree.write(dst, encoding="utf-8", xml_declaration=True)
        return True
    except Exception as exc:
        warnings.warn(f"格式化 XML 失败：{src}: {exc}")
        return False


def text_of(element):
    return "".join(t.text or "" for t in element.findall(".//w:t", NS))


def text_node_indexes(paragraph):
    return [i for i, node in enumerate(paragraph.findall(".//w:t", NS))]


def child_xpath(parent_path, tag_name, index):
    return f"{parent_path}/w:{tag_name}[{index + 1}]"


def parse_word_part(unpacked_dir, part, part_kind, start_table_index=0):
    path = unpacked_dir / part
    records = []
    paragraphs = []
    tables = []
    if not path.exists():
        return records, paragraphs, tables, start_table_index
    try:
        root = ET.parse(path).getroot()
    except Exception as exc:
        warnings.warn(f"解析 XML 失败：{part}: {exc}")
        return records, paragraphs, tables, start_table_index

    body = root.find("w:body", NS)
    container = body if body is not None else root
    paragraph_index = 0
    table_index = start_table_index
    part_prefix = "/w:document/w:body" if body is not None else f"/{root.tag.split('}', 1)[-1]}"

    for child_position, child in enumerate(list(container)):
        local = child.tag.split("}", 1)[-1]
        if local == "p":
            text = text_of(child).strip()
            if text:
                record_id = f"P{len(paragraphs) + 1:04d}" if part_kind == "body" else f"{part_kind.upper()}_P{len(paragraphs) + 1:04d}"
                xpath = child_xpath(part_prefix, "p", child_position)
                record = make_record(record_id, part, "paragraph" if part_kind == "body" else part_kind, text, xpath, paragraph_index, child)
                records.append(record)
                paragraphs.append({"id": record_id, "part": part, "index": paragraph_index, "text": text})
            paragraph_index += 1
        elif local == "tbl":
            table_id = f"T{table_index + 1:04d}"
            table_xpath = child_xpath(part_prefix, "tbl", sum(1 for previous in list(container)[:child_position] if previous.tag == qn("tbl")))
            table = {"id": table_id, "part": part, "index": table_index, "rows": []}
            rows = child.findall("w:tr", NS)
            for row_index, row in enumerate(rows):
                row_info = {"index": row_index, "cells": []}
                cells = row.findall("w:tc", NS)
                for cell_index, cell in enumerate(cells):
                    cell_id = f"{table_id}_R{row_index:03d}_C{cell_index:03d}"
                    cell_info = {"id": cell_id, "index": cell_index, "paragraphs": []}
                    cell_paragraphs = cell.findall("w:p", NS)
                    for cell_para_index, paragraph in enumerate(cell_paragraphs):
                        text = text_of(paragraph).strip()
                        if not text:
                            continue
                        record_id = f"{cell_id}_P{cell_para_index:03d}"
                        xpath = f"{table_xpath}/w:tr[{row_index + 1}]/w:tc[{cell_index + 1}]/w:p[{cell_para_index + 1}]"
                        record = make_record(record_id, part, "table_cell", text, xpath, cell_para_index, paragraph)
                        record.update({"table_index": table_index, "row_index": row_index, "cell_index": cell_index})
                        records.append(record)
                        cell_info["paragraphs"].append({"id": record_id, "text": text})
                    row_info["cells"].append(cell_info)
                table["rows"].append(row_info)
            tables.append(table)
            table_index += 1
    return records, paragraphs, tables, table_index


def make_record(record_id, part, kind, text, xpath, paragraph_index, paragraph):
    return {
        "id": record_id,
        "part": part,
        "kind": kind,
        "table_index": None,
        "row_index": None,
        "cell_index": None,
        "paragraph_index": paragraph_index,
        "run_index": [],
        "text_node_index": text_node_indexes(paragraph),
        "text": text,
        "xpath": xpath,
        "char_count": len(text),
        "contains_chinese": contains_chinese(text),
        "maybe_field": bool(LABEL_VALUE_RE.match(text) or PLACEHOLDER_RE.search(text)),
    }


def collect_structure(unpacked_dir):
    text_index = []
    paragraphs = []
    tables = []
    table_index = 0
    records, body_paragraphs, body_tables, table_index = parse_word_part(unpacked_dir, "word/document.xml", "body", table_index)
    text_index.extend(records)
    paragraphs.extend(body_paragraphs)
    tables.extend(body_tables)

    extra_parts = []
    for pattern, kind in [("word/header*.xml", "header"), ("word/footer*.xml", "footer")]:
        for path in sorted(unpacked_dir.glob(pattern)):
            extra_parts.append((path.relative_to(unpacked_dir).as_posix(), kind))
    for part, kind in [("word/footnotes.xml", "footnote"), ("word/endnotes.xml", "endnote")]:
        if (unpacked_dir / part).exists():
            extra_parts.append((part, kind))

    extras = {"headers": [], "footers": [], "footnotes": [], "endnotes": []}
    for part, kind in extra_parts:
        records, part_paragraphs, part_tables, table_index = parse_word_part(unpacked_dir, part, kind, table_index)
        text_index.extend(records)
        key = {"header": "headers", "footer": "footers", "footnote": "footnotes", "endnote": "endnotes"}[kind]
        extras[key].extend(part_paragraphs)
        tables.extend(part_tables)
    return {"paragraphs": paragraphs, "tables": tables, **extras}, text_index


def collect_text_nodes(unpacked_dir):
    nodes = []
    parts = ["word/document.xml"]
    parts.extend(path.relative_to(unpacked_dir).as_posix() for path in sorted(unpacked_dir.glob("word/header*.xml")))
    parts.extend(path.relative_to(unpacked_dir).as_posix() for path in sorted(unpacked_dir.glob("word/footer*.xml")))
    for optional in ["word/footnotes.xml", "word/endnotes.xml"]:
        if (unpacked_dir / optional).exists():
            parts.append(optional)
    for part in parts:
        path = unpacked_dir / part
        try:
            root = ET.parse(path).getroot()
        except Exception as exc:
            warnings.warn(f"收集文本节点失败：{part}: {exc}")
            continue
        for index, element in enumerate(root.findall(".//w:t", NS)):
            text = element.text or ""
            if text.strip():
                nodes.append({"global_index": len(nodes), "part": part, "part_text_index": index, "text": text.strip()})
    return nodes


def is_marker(text):
    return bool(MARKER_RE.match(text or ""))


def is_heading_text(text):
    stripped = (text or "").strip()
    if stripped in SECTION_TITLES:
        return True
    return 2 <= len(stripped) <= 8 and contains_chinese(stripped) and not any(ch in stripped for ch in "：:，,。；;（）()0123456789") and not is_marker(stripped)


def likely_content_direction(nodes, heading_index, window=8):
    previous_items = [node for node in nodes[max(0, heading_index - window):heading_index] if not is_marker(node["text"])]
    next_items = [node for node in nodes[heading_index + 1:heading_index + 1 + window] if not is_marker(node["text"])]
    previous_text = next((node["text"] for node in reversed(previous_items) if not is_heading_text(node["text"])), "")
    next_text = next((node["text"] for node in next_items if not is_heading_text(node["text"])), "")
    if previous_text and len(previous_text) >= len(next_text):
        return "before_heading"
    if next_text:
        return "after_heading"
    if previous_text:
        return "before_heading"
    return "unknown"


def write_section_context(path, text_nodes, window=10):
    headings = [node for node in text_nodes if is_heading_text(node["text"])]
    lines = [
        "# Section Context",
        "",
        "Use this file before editing DOCX XML. XML text order may not match Word visual layout, especially in resumes, text boxes, and duplicated drawing layers.",
        "",
    ]
    if not headings:
        lines.extend(["No likely section headings detected.", ""])
    for heading in headings:
        heading_index = heading["global_index"]
        lines.extend([
            f"## {heading['text']}",
            "",
            "Heading node:",
            f"- {heading_index}: {heading['text']} ({heading['part']} text_node {heading['part_text_index']})",
            "",
            f"Likely content direction: {likely_content_direction(text_nodes, heading_index)}",
            "",
            "Nearby previous nodes:",
        ])
        for node in text_nodes[max(0, heading_index - window):heading_index]:
            lines.append(f"- {node['global_index']}: {node['text']}")
        lines.extend(["", "Nearby next nodes:"])
        for node in text_nodes[heading_index + 1:heading_index + 1 + window]:
            lines.append(f"- {node['global_index']}: {node['text']}")
        lines.extend([
            "",
            "Warning:",
            "- Verify section ownership before editing; do not replace text only because it matches a phrase.",
            "- If the same heading appears twice, the document may contain duplicated visual/text layers; edit matching duplicated content nodes together.",
            "",
        ])
    path.write_text("\n".join(lines), encoding="utf-8")


def write_structure(path, source_name, manifest, structure, text_index):
    xml_count = sum(1 for item in manifest if item["type"] in {"xml", "rels"})
    media_count = sum(1 for item in manifest if item["type"] == "media")
    lines = [
        "# DOCX Structure",
        "",
        f"Source: {source_name}",
        "",
        "## Summary",
        "",
        f"- Parts: {len(manifest)}",
        f"- XML files: {xml_count}",
        f"- Media files: {media_count}",
        f"- Paragraphs: {len(structure['paragraphs'])}",
        f"- Tables: {len(structure['tables'])}",
        f"- Text records: {len(text_index)}",
        "",
    ]
    if structure["paragraphs"]:
        lines.extend(["## Body Paragraphs", ""])
        for item in structure["paragraphs"]:
            lines.extend([f"[{item['id']}] {item['part']} paragraph {item['index']}", item["text"], ""])
    if structure["tables"]:
        lines.extend(["## Tables", ""])
        for table in structure["tables"]:
            lines.extend([f"### Table {table['id']}", ""])
            for row in table["rows"]:
                lines.append(f"Row {row['index']}:")
                for cell in row["cells"]:
                    lines.append(f"- Cell {cell['index']} [{cell['id']}]:")
                    for paragraph in cell["paragraphs"]:
                        lines.append(f"  - {paragraph['id'].split('_')[-1]}: {paragraph['text']}")
                lines.append("")
    for title, key in [("Headers", "headers"), ("Footers", "footers"), ("Footnotes", "footnotes"), ("Endnotes", "endnotes")]:
        if structure[key]:
            lines.extend([f"## {title}", ""])
            for item in structure[key]:
                lines.extend([f"[{item['id']}] {item['part']} paragraph {item['index']}", item["text"], ""])
    path.write_text("\n".join(lines), encoding="utf-8")


def yaml_quote(value):
    text = str(value).replace('"', '\\"')
    return f'"{text}"'


def write_candidate_fields(path, text_index):
    fields = []
    for item in text_index:
        text = item["text"].strip()
        match = LABEL_VALUE_RE.match(text)
        if match:
            label, value = match.group(1), match.group(2)
            fields.append({
                "key": normalize_key(label),
                "label": label,
                "current_text": value,
                "confidence": 0.9,
                "reason": "text contains label-value pattern",
                "location": location_from_item(item),
            })
        elif PLACEHOLDER_RE.search(text):
            fields.append({
                "key": f"field_{len(fields) + 1}",
                "label": "待填写字段",
                "current_text": text,
                "confidence": 0.75,
                "reason": "text contains placeholder marker",
                "location": location_from_item(item),
            })
    lines = ["fields:"]
    if not fields:
        lines.append("  []")
    for field in fields:
        lines.append(f"  - key: {yaml_quote(field['key'])}")
        lines.append(f"    label: {yaml_quote(field['label'])}")
        lines.append(f"    current_text: {yaml_quote(field['current_text'])}")
        lines.append(f"    confidence: {field['confidence']}")
        lines.append(f"    reason: {yaml_quote(field['reason'])}")
        lines.append("    location:")
        for key, value in field["location"].items():
            if value is not None:
                lines.append(f"      {key}: {yaml_quote(value) if isinstance(value, str) else value}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return fields


def normalize_key(label):
    mapping = {"姓名": "name", "学号": "student_id", "专业": "major", "日期": "date", "题目": "title"}
    return mapping.get(label, re.sub(r"\W+", "_", label).strip("_") or "field")


def location_from_item(item):
    return {
        "part": item["part"],
        "id": item["id"],
        "table_index": item.get("table_index"),
        "row_index": item.get("row_index"),
        "cell_index": item.get("cell_index"),
        "paragraph_index": item.get("paragraph_index"),
        "xpath": item.get("xpath"),
    }


def safe_workspace_name(input_path):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    stem = re.sub(r'[<>:"/\\|?*\s]+', "_", input_path.stem).strip("_") or "docx"
    return f"{timestamp}_{stem}"


def append_workspace_index(workspace_root, source_name, output_path):
    index_path = workspace_root / "index.md"
    if not index_path.exists():
        index_path.write_text("# DOCX XML Workspaces\n\n| Time | Source | Workspace |\n|---|---|---|\n", encoding="utf-8")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with index_path.open("a", encoding="utf-8") as f:
        f.write(f"| {now} | {source_name} | {output_path.name} |\n")


def export_docx_xml(input_docx, output_dir=None, workspace_root=None):
    input_path = Path(input_docx).resolve()
    validate_docx(input_path)
    if output_dir:
        output_path = Path(output_dir).resolve()
    elif workspace_root:
        workspace_root = Path(workspace_root).resolve()
        workspace_root.mkdir(parents=True, exist_ok=True)
        output_path = workspace_root / safe_workspace_name(input_path)
    else:
        output_path = input_path.parent / f"{input_path.stem}_docx_xml"
    unpacked_dir = output_path / "unpacked"
    pretty_dir = output_path / "pretty_xml"

    if output_path.exists():
        shutil.rmtree(output_path)
    unpacked_dir.mkdir(parents=True)
    pretty_dir.mkdir(parents=True)
    shutil.copy2(input_path, output_path / "original.docx")

    with zipfile.ZipFile(input_path) as z:
        z.extractall(unpacked_dir)

    content_types = read_content_types(unpacked_dir)
    manifest = []
    for file in sorted(path for path in unpacked_dir.rglob("*") if path.is_file()):
        rel_path = file.relative_to(unpacked_dir).as_posix()
        kind = file_type(rel_path)
        pretty_printed = kind in {"xml", "rels"} and write_pretty_xml(file, pretty_dir / rel_path)
        manifest.append({
            "path": rel_path,
            "type": kind,
            "size": file.stat().st_size,
            "whether_pretty_printed": pretty_printed,
            "content_type": content_type_for(rel_path, content_types),
        })

    structure, text_index = collect_structure(unpacked_dir)
    text_nodes = collect_text_nodes(unpacked_dir)
    (output_path / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_path / "text_index.json").write_text(json.dumps(text_index, ensure_ascii=False, indent=2), encoding="utf-8")
    write_structure(output_path / "structure.md", input_path.name, manifest, structure, text_index)
    write_section_context(output_path / "section_context.md", text_nodes)
    fields = write_candidate_fields(output_path / "candidate_fields.yaml", text_index)
    if workspace_root and not output_dir:
        append_workspace_index(Path(workspace_root).resolve(), input_path.name, output_path)

    print_summary(output_path, manifest, structure, text_index, fields)
    return output_path


def next_edit_paths(workspace_dir, output_docx=None):
    edits_dir = workspace_dir / "edits"
    edits_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(edits_dir.glob("output_*.docx"))
    next_number = len(existing) + 1
    audit_path = edits_dir / f"edit_audit_{next_number:03d}.md"
    if output_docx:
        output_path = Path(output_docx).resolve()
    else:
        output_path = edits_dir / f"output_{next_number:03d}.docx"
    return edits_dir, output_path, audit_path


def pack_docx_from_directory(source_dir, output_docx):
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    if output_docx.exists():
        output_docx.unlink()
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(source_dir).as_posix())
    with zipfile.ZipFile(output_docx) as z:
        bad = z.testzip()
    return bad or "ok"


def fill_docx_workspace(workspace_dir, replacements_json, output_docx=None):
    workspace_dir = Path(workspace_dir).resolve()
    replacements_json = Path(replacements_json).resolve()
    unpacked_dir = workspace_dir / "unpacked"
    if not unpacked_dir.exists():
        raise FileNotFoundError(f"找不到工作区 unpacked 目录：{unpacked_dir}")
    with replacements_json.open("r", encoding="utf-8") as f:
        spec = json.load(f)
    replacements = spec.get("replacements", [])
    if not replacements:
        raise ValueError("replacements.json 中没有 replacements")

    edits_dir, output_path, audit_path = next_edit_paths(workspace_dir, output_docx)
    work_dir = edits_dir / "working_unpacked"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    shutil.copytree(unpacked_dir, work_dir)

    audit_lines = ["# Edit Audit", "", f"Workspace: {workspace_dir}", f"Replacements file: {replacements_json}", ""]
    parsed_trees = {}
    for replacement in replacements:
        label = replacement.get("label", "")
        part = replacement["part"]
        indexes = replacement["text_node_indexes"]
        old_text = replacement["old_text"]
        new_text = replacement["new_text"]
        part_path = work_dir / part
        if not part_path.exists():
            raise FileNotFoundError(f"找不到 XML part：{part}")
        if part not in parsed_trees:
            parsed_trees[part] = ET.parse(part_path)
        tree = parsed_trees[part]
        text_nodes = tree.getroot().findall(".//w:t", NS)
        audit_lines.extend([f"## {label or part}", "", f"label: {label}", f"part: {part}", f"text_node_indexes: {indexes}"])
        for index in indexes:
            if index < 0 or index >= len(text_nodes):
                raise IndexError(f"text_node index 超出范围：{part} #{index}")
            before = text_nodes[index].text or ""
            matched = before == old_text
            audit_lines.extend([f"- node {index}", f"  before: {before}", f"  old_text matched: {matched}"])
            if not matched:
                raise ValueError(f"old_text 不匹配：{part} #{index}，实际为 {before!r}")
            text_nodes[index].text = new_text
            audit_lines.append(f"  after: {new_text}")
        audit_lines.append("")

    for part, tree in parsed_trees.items():
        tree.write(work_dir / part, encoding="utf-8", xml_declaration=True)

    zip_test = pack_docx_from_directory(work_dir, output_path)
    audit_lines.extend(["## Verification", "", f"output_docx: {output_path}", f"zip_test: {zip_test}", ""])
    audit_path.write_text("\n".join(audit_lines), encoding="utf-8")
    return output_path
def print_summary(output_path, manifest, structure, text_index, fields):
    print("导出完成")
    print(f"输出目录：{output_path}")
    print(f"XML/关系文件：{sum(1 for item in manifest if item['type'] in {'xml', 'rels'})}")
    print(f"媒体文件：{sum(1 for item in manifest if item['type'] == 'media')}")
    print(f"段落数量：{len(structure['paragraphs'])}")
    print(f"表格数量：{len(structure['tables'])}")
    print(f"文本索引数量：{len(text_index)}")
    print(f"候选字段数量：{len(fields)}")


def self_test():
    print("Python:", sys.version.split()[0])
    print("zipfile: ok")
    print("xml.etree.ElementTree: ok")
    print("工具自检通过")


def main():
    parser = argparse.ArgumentParser(description="DOCX XML 导出与结构索引工具")
    parser.add_argument("input", nargs="?", help="输入 .docx 文件路径")
    parser.add_argument("--out", help="输出目录；指定后不会写入工作区 index.md")
    parser.add_argument("--workspace-root", help="统一工作区根目录；未指定 --out 时输出到 时间_文件名 子目录并更新 index.md")
    parser.add_argument("--test", action="store_true", help="运行自检")
    args = parser.parse_args()
    if args.test:
        self_test()
        return
    if not args.input:
        parser.print_help()
        return
    export_docx_xml(args.input, args.out, args.workspace_root)


if __name__ == "__main__":
    main()
