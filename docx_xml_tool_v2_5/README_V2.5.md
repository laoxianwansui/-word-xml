# DOCX XML Tool v2.5 Overview

v2.5 adds a local Web GUI on top of the v2.4.2 DOCX XML preservation core. It does not rebuild Word files from scratch and does not modify the original DOCX.

## What changed from v2.4.2

| Area | v2.4.2 | v2.5 |
|---|---|---|
| Core export | CLI and drag-and-drop workspace export | Same core |
| Markdown apply | CLI and drag-and-drop `document.md` writeback | Same core |
| Field fill | CLI and drag-and-drop `field_values.yaml` writeback | Same core |
| User interface | CLI and three Windows launchers | Adds local browser GUI and `04_start_web_gui.bat` |
| Network model | Local files only | Local HTTP server on `127.0.0.1` by default |

## Recommended GUI workflow

```text
Run launchers/04_start_web_gui.bat
  ↓
Browser opens http://127.0.0.1:8765/
  ↓
Use 01 Export Workspace for a .docx path
  ↓
Edit document.md or field_values.yaml in the generated workspace
  ↓
Use 02 Apply Markdown or 03 Fill Fields
  ↓
Open the generated DOCX under edits/ and visually inspect it
```

## CLI equivalents

```bash
python docx_workspace_tool.py --web-gui
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## Current files

```text
docx_workspace_tool.py
docx_web_gui.py
docx_markdown/engine.py
docx_markdown/fields.py
launchers/01_export_docx_workspace.bat
launchers/02_apply_markdown_to_docx.bat
launchers/03_fill_fields_to_docx.bat
launchers/04_start_web_gui.bat
scripts/package_release.py
test_docx_workspace_tool.py
```

## Boundaries

v2.5 intentionally does not add cloud sync, login, AI editing, MCP, multi-document RAG, OCR, PDF conversion, or full Word layout reconstruction. Those remain outside this release.

## Packaging

The release zip is:

```text
dist/docx_xml_tool_v2_5.zip
```

Generated workspaces, logs, caches, `edits/`, and private documents must not be included.
