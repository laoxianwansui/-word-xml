# DOCX XML Tool Maintenance Guide

This file is the first place to read when continuing development after a long pause or after Claude context compression.

## Workspace locations

Repository root:

```text
F:\word-xml-work
```

Current package:

```text
F:\word-xml-work\docx_xml_tool_v2_4_2
```

Historical versions:

```text
F:\word-xml-work\archive
```

Work from the F drive repo, not from Desktop copies.

## Read these after context compression

If the conversation context is compressed or stale, read these files first:

```text
README_MAINTENANCE.md
README.md
README.zh-CN.md
README_V2.4.2.md
..\README.md
..\archive\README.md
```

Then run `git status --short --branch` before editing.

## Version iteration rule

Every version iteration must use matching names. Do not put v2.4.3 code in a v2.4.2 folder.

For example, when moving from v2.4.2 to v2.4.3:

1. Archive the previous current package:

   ```text
   docx_xml_tool_v2_4_2 -> archive/docx_xml_tool_v2_4_2
   ```

2. Create the new current package:

   ```text
   docx_xml_tool_v2_4_3
   ```

3. Rename the version overview file:

   ```text
   README_V2.4.2.md -> README_V2.4.3.md
   ```

4. Update all current-version references:

   ```text
   README.md
   README.zh-CN.md
   README_V2.4.3.md
   ..\README.md
   ..\archive\README.md
   scripts\package_release.py
   ```

5. Grep before reporting completion:

   ```bash
   grep for: docx_xml_tool_v2_4_2
   grep for: README_V2.4.2
   grep for: docx_xml_tool_v2_4_2.zip
   ```

Historical mentions are allowed only in archive docs or comparison text. Current package, scripts, tests, and release names must use the current version.

## Current scripts

Windows drag-and-drop launchers:

```text
launchers\01_export_docx_workspace.bat      # export a DOCX workspace
launchers\02_apply_markdown_to_docx.bat     # apply document.md back to DOCX
launchers\03_fill_fields_to_docx.bat        # apply field_values.yaml back to DOCX
```

CLI equivalents:

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## Test and package commands

Run from the repo root or use absolute paths.

Unit tests:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_4_2/test_docx_workspace_tool.py"
```

Self-test:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_4_2/docx_workspace_tool.py" --test
```

Create release zip for inspection:

```bash
python "F:/word-xml-work/docx_xml_tool_v2_4_2/scripts/package_release.py"
```

Expected release zip:

```text
F:\word-xml-work\docx_xml_tool_v2_4_2\dist\docx_xml_tool_v2_4_2.zip
```

Delete `dist/` after inspection unless the user explicitly asks to keep the zip.

## Process artifact cleanup

Before reporting completion, delete generated/process files:

```text
dist/
__pycache__/
.pytest_cache/
docx_xml_outputs/
*_docx_xml/
edits/
*.log
docs/superpowers/
.claude plan files created only for this task
```

Do not delete source files, README files, launchers, tests, scripts, or archived release folders.

## Archive rule

Each new iteration should archive the previous current package under:

```text
archive/docx_xml_tool_vX_Y_Z
```

Only one current package should remain at the repository root. Older versions in `archive/` are for reference and rollback.

Update `archive/README.md` whenever a version is archived.

## Current v2.4.2 feature summary

v2.4.2 adds Smart Field Fill on top of the v2.4.1-safe Markdown roundtrip:

```text
fields.yaml
field_values.yaml
--fill-fields
field_fill_output_###.docx
field_fill_audit_###.md
```

It also detects table-adjacent fields and marks content controls as high-risk field locations.

## GitHub rule

Do not push to GitHub unless the user explicitly asks. Do not commit unless the user explicitly asks.
