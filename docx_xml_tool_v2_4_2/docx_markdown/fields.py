from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List

LABEL_VALUE_RE = re.compile(r"^\s*([一-鿿A-Za-z0-9_（）()]{1,12})\s*[:：]\s*(.*?)\s*$")
PLACEHOLDER_RE = re.compile(r"(XXX+|_{2,}|—{2,}|请填写|请输入|待填写)")
KEY_MAP = {"姓名": "name", "学号": "student_id", "专业": "major", "日期": "date", "题目": "title"}


def normalize_key(label: str) -> str:
    if label in KEY_MAP:
        return KEY_MAP[label]
    key = re.sub(r"\W+", "_", label).strip("_").lower()
    return key or "field"


def detect_fields_from_blocks(blocks: Iterable[Dict[str, object]]) -> List[Dict[str, object]]:
    block_list = list(blocks)
    fields: List[Dict[str, object]] = []
    seen_keys: Dict[str, int] = {}
    for block in block_list:
        text = str(block.get("text", "")).strip()
        if not text:
            continue
        match = LABEL_VALUE_RE.match(text)
        if match:
            label, value = match.group(1), match.group(2)
            base_key = normalize_key(label)
            key = unique_key(base_key, seen_keys)
            fields.append(make_field(key, label, value, "replace_value", 0.9, risk_for_block(block, "low"), block))
            continue
        if PLACEHOLDER_RE.search(text):
            base_key = f"field_{len(fields) + 1}"
            key = unique_key(base_key, seen_keys)
            fields.append(make_field(key, "待填写字段", text, "replace_all", 0.75, risk_for_block(block, "medium"), block))
    fields.extend(detect_table_adjacent_fields(block_list, seen_keys))
    return fields


def detect_table_adjacent_fields(blocks: List[Dict[str, object]], seen_keys: Dict[str, int]) -> List[Dict[str, object]]:
    fields: List[Dict[str, object]] = []
    by_table_row: Dict[tuple, List[Dict[str, object]]] = {}
    for block in blocks:
        if block.get("kind") != "table_cell":
            continue
        table_index = block.get("table_index")
        row_index = block.get("row_index")
        if not isinstance(table_index, int) or not isinstance(row_index, int):
            continue
        by_table_row.setdefault((str(block.get("part", "")), table_index, row_index), []).append(block)
    for row_blocks in by_table_row.values():
        ordered = sorted(row_blocks, key=lambda item: int(item.get("col_index", 0) or 0))
        for left, right in zip(ordered, ordered[1:]):
            label = str(left.get("text", "")).strip().rstrip("：:")
            value = str(right.get("text", "")).strip()
            if not is_short_label(label) or not value:
                continue
            base_key = normalize_key(label)
            key = unique_key(base_key, seen_keys)
            fields.append(make_field(key, label, value, "replace_all", 0.8, risk_for_block(right, "low"), right))
    return fields


def is_short_label(text: str) -> bool:
    return bool(text) and len(text) <= 12 and not PLACEHOLDER_RE.search(text) and not LABEL_VALUE_RE.match(text)


def risk_for_block(block: Dict[str, object], default: str) -> str:
    if block.get("kind") in {"content_control", "textbox"}:
        return "high"
    return default


def unique_key(base_key: str, seen_keys: Dict[str, int]) -> str:
    count = seen_keys.get(base_key, 0) + 1
    seen_keys[base_key] = count
    if count == 1:
        return base_key
    return f"{base_key}_{count}"


def make_field(key: str, label: str, current_text: str, replacement_mode: str, confidence: float, risk: str, block: Dict[str, object]) -> Dict[str, object]:
    return {
        "key": key,
        "label": label,
        "current_text": current_text,
        "replacement_mode": replacement_mode,
        "confidence": confidence,
        "risk": risk,
        "locations": [{
            "block_id": str(block.get("block_id", "")),
            "part": str(block.get("part", "")),
            "kind": str(block.get("kind", "")),
            "path": str(block.get("path", "")),
        }],
    }


def write_field_files(workspace_dir: Path, content_map: Dict[str, object]) -> List[Dict[str, object]]:
    fields = detect_fields_from_blocks(content_map.get("blocks", []))
    workspace_dir.mkdir(parents=True, exist_ok=True)
    (workspace_dir / "fields.yaml").write_text(render_fields_yaml(fields), encoding="utf-8")
    (workspace_dir / "field_values.yaml").write_text(render_field_values_yaml(fields), encoding="utf-8")
    return fields


def render_fields_yaml(fields: List[Dict[str, object]]) -> str:
    lines = ["fields:"]
    if not fields:
        lines.append("  []")
        return "\n".join(lines) + "\n"
    for field in fields:
        lines.extend([
            f"  - key: {field['key']}",
            f"    label: {field['label']}",
            f"    current_text: {field['current_text']}",
            f"    replacement_mode: {field['replacement_mode']}",
            f"    confidence: {field['confidence']}",
            f"    risk: {field['risk']}",
            "    locations:",
        ])
        for location in field["locations"]:
            lines.extend([
                f"      - block_id: {location['block_id']}",
                f"        part: {location['part']}",
                f"        kind: {location['kind']}",
                f"        path: {location['path']}",
            ])
    return "\n".join(lines) + "\n"


def render_field_values_yaml(fields: List[Dict[str, object]]) -> str:
    lines = ["# Edit values after the colon, then run --fill-fields."]
    for field in fields:
        lines.append(f"# {field['label']} | current: {field['current_text']}")
        lines.append(f"{field['key']}: ")
    return "\n".join(lines) + "\n"


def parse_field_values(text: str) -> Dict[str, str]:
    values: Dict[str, str] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        values[key.strip()] = value.strip()
    return values


def plan_field_replacements(fields: List[Dict[str, object]], values: Dict[str, str], blocks: List[Dict[str, object]]) -> Dict[str, object]:
    block_by_id = {str(block.get("block_id", "")): block for block in blocks}
    replacements: Dict[str, str] = {}
    changed_fields: List[Dict[str, object]] = []
    skipped_fields: List[Dict[str, object]] = []
    warnings: List[str] = []

    for field in fields:
        key = str(field.get("key", ""))
        if key not in values or values[key] == "":
            continue
        locations = field.get("locations", [])
        if not isinstance(locations, list) or len(locations) != 1:
            warnings.append(f"ambiguous field skipped: {key}")
            skipped_fields.append({"key": key, "reason": "ambiguous locations"})
            continue
        location = locations[0]
        block_id = str(location.get("block_id", ""))
        block = block_by_id.get(block_id)
        if not block:
            warnings.append(f"field skipped because block is missing: {key}")
            skipped_fields.append({"key": key, "reason": "missing block"})
            continue
        old_text = str(block.get("text", ""))
        new_value = values[key]
        replacement_mode = str(field.get("replacement_mode", ""))
        label = str(field.get("label", ""))
        if replacement_mode == "replace_value":
            new_text = replace_label_value(old_text, label, new_value)
        else:
            new_text = new_value
        if old_text != new_text:
            replacements[block_id] = new_text
            changed_fields.append({"key": key, "block_id": block_id, "old_text": old_text, "new_text": new_text})
    return {"replacements": replacements, "changed_fields": changed_fields, "skipped_fields": skipped_fields, "warnings": warnings}


def replace_label_value(old_text: str, label: str, new_value: str) -> str:
    match = LABEL_VALUE_RE.match(old_text)
    if not match:
        return new_value
    actual_label = match.group(1)
    separator_match = re.search(r"[:：]", old_text)
    separator = separator_match.group(0) if separator_match else "："
    return f"{actual_label}{separator}{new_value}"
