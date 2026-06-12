# DOCX XML Tool Maintenance Guide

This file is the first place to read when continuing development after a long pause or after Claude context compression.

## Workspace locations

Repository root:

```text
F:\word-xml-work
```

Current package:

```text
F:\word-xml-work\docx_xml_tool_v2_5
```

Historical versions:

```text
F:\word-xml-work\archive
```

Work from the F drive repo, not from Desktop copies.

## Read these after context compression

```text
README_MAINTENANCE.md
README.md
README.zh-CN.md
README_V2.5.md
..\README.md
..\archive\README.md
```

Then run `git status --short --branch` before editing.

## Current v2.5 feature summary

v2.5 adds a local Web GUI on top of the v2.4.2-safe XML/Markdown/field-fill core:

```text
docx_web_gui.py
--web-gui
--host
--port
--no-open
launchers\04_start_web_gui.bat
```

The GUI is local-only by default and listens on `127.0.0.1:8765` unless changed by CLI arguments.

## Current scripts

Windows launchers:

```text
launchers\01_export_docx_workspace.bat      # export a DOCX workspace
launchers\02_apply_markdown_to_docx.bat     # apply document.md back to DOCX
launchers\03_fill_fields_to_docx.bat        # apply field_values.yaml back to DOCX
launchers\04_start_web_gui.bat              # start local Web GUI
```

CLI equivalents:

```bash
python docx_workspace_tool.py --web-gui
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## Test and package commands

Unit tests:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_5/test_docx_workspace_tool.py"
```

Self-test:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_5/docx_workspace_tool.py" --test
```

Create release zip for inspection:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_5/scripts/package_release.py"
```

Expected release zip:

```text
F:\word-xml-work\docx_xml_tool_v2_5\dist\docx_xml_tool_v2_5.zip
```

Delete `dist/` after inspection unless the user explicitly asks to keep the zip.

## Version iteration rule

Every version iteration must use matching names. Do not put one version's code in another version's folder.

When moving to a future version:

1. Archive the previous current package under `archive/docx_xml_tool_vX_Y_Z`.
2. Create the new current package at the repository root.
3. Rename the version overview file to match the new version.
4. Update `README.md`, `README.zh-CN.md`, the version overview, `..\README.md`, `..\archive\README.md`, and `scripts\package_release.py`.
5. Grep for old current-version folder names, old README names, and old zip names before reporting completion.

Historical mentions are allowed only in archive docs or comparison text. Current package, scripts, tests, and release names must use the current version.

## Process artifact cleanup

Before reporting completion, delete generated/process files unless the user asked to keep them:

```text
dist/
__pycache__/
docx_markdown/__pycache__/
.pytest_cache/
docx_xml_outputs/
*_docx_xml/
edits/
*.log
docs/superpowers/
```

Do not delete source files, README files, launchers, tests, scripts, or archived release folders.

## Archive rule

Each new iteration should archive the previous current package under:

```text
archive/docx_xml_tool_vX_Y_Z
```

Only one current package should remain at the repository root. Older versions in `archive/` are for reference and rollback.

Update `archive/README.md` whenever a version is archived.

## GitHub rule

Do not push to GitHub unless the user explicitly asks. Do not commit unless the user explicitly asks.
