# DOCX XML Tool

A small, dependency-free Python tool for unpacking `.docx` files into an AI-friendly XML workspace, indexing document text, and safely writing selected text changes back into a new Word document.

[中文说明](README.zh-CN.md)

## Why this exists

`.docx` files are ZIP archives containing XML. Editing them through high-level Word libraries can accidentally rebuild paragraphs, lose run-level formatting, or change layout. This tool keeps the original DOCX structure intact:

1. Copy the original file.
2. Unpack the full DOCX archive.
3. Pretty-print XML for reading.
4. Generate structure and text indexes for AI or manual inspection.
5. Optionally create a new DOCX by replacing specific XML text nodes.

The original `.docx` file is never modified.

## Features

- Export a DOCX into a workspace with `original.docx`, `unpacked/`, and `pretty_xml/`.
- Generate `structure.md` for quick document/table inspection.
- Generate `section_context.md` to reduce mistakes when nearby visual order differs from XML order.
- Generate `text_index.json` with table, row, cell, paragraph, XPath, and text metadata.
- Generate `candidate_fields.yaml` for likely label-value fields and placeholders.
- Fill selected text nodes back into a new DOCX while preserving the original ZIP/XML structure.
- Write an edit audit file for each generated output.
- Supports Chinese paths, Chinese filenames, and Chinese document text when run with a UTF-8 capable environment.

## Requirements

- Python 3.9 or newer
- No third-party Python packages

## Quick start

```bash
python docx_xml_exporter.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

The output directory will look like:

```text
docx_xml_outputs/
  index.md
  2026-06-09_153000_input/
    original.docx
    unpacked/
    pretty_xml/
    manifest.json
    structure.md
    section_context.md
    text_index.json
    candidate_fields.yaml
```

You can also choose a fixed output directory:

```bash
python docx_xml_exporter.py "C:/path/to/input.docx" --out "C:/path/to/workspace"
```

When `--out` is used, the workspace index is not updated.

## Windows drag-and-drop launcher

A Windows launcher is included at:

```text
launchers/DOCX_XML_export.bat
```

After cloning or downloading the project, drag a `.docx` file onto this `.bat` file, or double-click it and paste a full `.docx` path.

The launcher writes outputs to:

```text
docx_xml_outputs/
```

and logs the latest run to:

```text
docx_xml_outputs/last_run.log
```

The launcher avoids common Windows batch pitfalls with Chinese paths and parentheses in filenames.

## Workspace files

| File or folder | Purpose |
|---|---|
| `original.docx` | A copy of the source DOCX. |
| `unpacked/` | The raw DOCX ZIP contents. Use this for safe XML-based editing. |
| `pretty_xml/` | Pretty-printed XML for reading only. Do not write this back. |
| `manifest.json` | File list, sizes, content types, and XML pretty-print status. |
| `structure.md` | Human-readable document structure: body paragraphs, tables, headers, footers, footnotes, and endnotes. |
| `section_context.md` | Nearby text around likely headings, useful before editing. |
| `text_index.json` | Machine-readable text records with table/cell/paragraph metadata. |
| `candidate_fields.yaml` | Heuristic list of likely fields and placeholders. |
| `edits/` | Generated DOCX outputs and edit audit files. |

## Filling text back into DOCX

`fill_docx_workspace()` can replace explicit XML text nodes and create a new `.docx` without modifying the original workspace.

Example replacement spec:

```json
{
  "replacements": [
    {
      "label": "Student name",
      "part": "word/document.xml",
      "text_node_indexes": [12],
      "old_text": "Name: Zhang San",
      "new_text": "Name: Li Si"
    }
  ]
}
```

Use it from Python:

```python
from pathlib import Path
from docx_xml_exporter import fill_docx_workspace

workspace = Path("C:/path/to/workspace")
replacements = workspace / "replacements.json"
output_docx = workspace / "edited.docx"

fill_docx_workspace(workspace, replacements, output_docx)
```

The function refuses to write when `old_text` does not match the target node. This prevents accidental edits to the wrong XML location.

## Testing

Run the self-test:

```bash
python docx_xml_exporter.py --test
```

Run the unit tests:

```bash
python test_docx_xml_exporter.py
```

## Safety notes

- The tool does not upload documents.
- The tool does not use network access.
- The source DOCX is not modified.
- Generated DOCX files are written under `edits/` or the output path you provide.
- `pretty_xml/` is for reading only; write back from `unpacked/`.

## Limitations

- This is not a full Word rendering engine.
- XML text order may differ from Word visual order, especially with text boxes, drawing layers, headers, footers, and duplicated content.
- The current release focuses on DOCX XML export and explicit text-node replacement.
- Markdown knowledge-base export is not implemented yet.

## Roadmap

- Export `content.md` for knowledge-base ingestion.
- Generate `content_map.json` linking Markdown rows/sections back to DOCX XML nodes.
- Support Markdown-to-DOCX XML content sync while preserving the original Word layout.
- Add a more ergonomic CLI subcommand for fill operations.

## License

MIT
