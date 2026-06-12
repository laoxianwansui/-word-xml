# DOCX XML Tool v2.5

A dependency-free local Python tool for unpacking `.docx` files into an XML/Markdown workspace, safely writing selected text changes back into a new Word document, and running the workflow from a local Web GUI.

[中文说明](README.zh-CN.md)

## What v2.5 adds

v2.5 keeps the v2.4.2 XML-preserving core and adds a local browser workbench. The GUI runs on your own computer at `127.0.0.1`; documents are not uploaded.

The GUI wraps the three stable workflows:

1. Export a `.docx` into a workspace.
2. Apply edited `document.md` back to a new `.docx`.
3. Fill detected fields from `field_values.yaml` into a new `.docx`.

## Quick start

### Local drag-drop Web GUI

Double-click:

```text
launchers/04_start_web_gui.bat
```

The browser opens:

```text
http://127.0.0.1:8765/
```

Drag a `.docx` file into the large drop zone. The tool automatically:

1. saves the uploaded file under `gui_uploads/`;
2. exports a workspace under `docx_xml_outputs/`;
3. shows buttons to open `document.md`, open `field_values.yaml`, apply Markdown, fill fields, and open the output folder.

Generated Word files are written under the workspace:

```text
edits/markdown_output_001.docx
edits/field_fill_output_001.docx
```

Command line GUI startup is still available:

```bash
python docx_workspace_tool.py --web-gui
python docx_workspace_tool.py --web-gui --no-open
```

### Backup CLI commands

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

### CLI Markdown apply

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

### CLI field fill

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## Workflow

```text
DOCX input
  ↓
Export workspace
  ↓
original.docx + unpacked/ + pretty_xml/
  ↓
structure.md + section_context.md + text_index.json
  ↓
document.md + content_map.json + fields.yaml + field_values.yaml
  ↓
Edit document.md or field_values.yaml
  ↓
Apply through CLI, launcher, or local Web GUI
  ↓
edits/*.docx + audit files
```

The original `.docx` is never modified.

## Windows launchers

```text
launchers/01_export_docx_workspace.bat      # drag a .docx file here to export a workspace
launchers/02_apply_markdown_to_docx.bat     # drag an exported workspace folder here to apply document.md
launchers/03_fill_fields_to_docx.bat        # drag an exported workspace folder here to apply field_values.yaml
launchers/04_start_web_gui.bat              # start the local browser GUI
```

`03_fill_fields_to_docx.bat` expects an exported workspace folder, not a raw `.docx` file. Run `01_export_docx_workspace.bat` first, edit `field_values.yaml`, then drag the workspace folder onto `03_fill_fields_to_docx.bat`.

## Workspace files

| File or folder | Purpose |
|---|---|
| `original.docx` | Copy of the source DOCX. |
| `unpacked/` | Raw DOCX ZIP contents used for preservation. |
| `pretty_xml/` | Pretty-printed XML for reading only. |
| `manifest.json` | File list, sizes, content types, and XML status. |
| `structure.md` | Human-readable document structure. |
| `section_context.md` | Nearby heading/context hints for safer editing. |
| `text_index.json` | Machine-readable XML text records. |
| `candidate_fields.yaml` | Legacy candidate field list. |
| `document.md` | Markdown editing surface. |
| `content_map.json` | Block-to-DOCX XML mapping required for writeback. |
| `format_profile.yaml` | Preservation policy summary. |
| `fields.yaml` | Detected field-fill candidates. |
| `field_values.yaml` | User-editable field values for `--fill-fields`. |
| `edits/` | Generated DOCX outputs and audit files. |

## Safety model

- Preserve the original DOCX package and rewrite only mapped XML parts.
- Refuse Markdown writeback when block markers are missing, duplicated, or unknown.
- Keep package relationships, media, headers, footers, footnotes, and endnotes unless their mapped text is edited.
- Record changed blocks in audit files for manual inspection.
- Require opening important generated DOCX files in Word/WPS/LibreOffice before final use.

## Not included in v2.5

- Cloud upload, login, collaboration, or remote storage.
- MCP server integration.
- AI editing agents.
- Multi-document RAG or vector search.
- Full Word rendering or layout reconstruction.
- Editing images, charts, SmartArt, equations, comments, tracked changes, macros, or embedded files.

## Naming and packaging

- Release folder name: `docx_xml_tool` or `docx_xml_tool_v2_5`.
- Release zip name: `dist/docx_xml_tool_v2_5.zip`.
- Do not package generated `docx_xml_outputs/`, `*_docx_xml/`, `edits/`, `__pycache__/`, `.pytest_cache/`, logs, or private documents.

Create a clean release zip:

```bash
python scripts/package_release.py
```

## Testing

```bash
python docx_workspace_tool.py --test
python test_docx_workspace_tool.py
```

## Requirements

- Python 3.9 or newer. Python 3.13 is tested.
- No third-party Python packages are required.
- CLI works on Windows, macOS, and Linux.
- `.bat` launchers are Windows-only.

## License

MIT
