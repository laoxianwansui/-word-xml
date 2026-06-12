# Archived versions

This folder keeps older DOCX XML Tool releases for reference and rollback.

| Version | Folder | Status | Purpose |
|---|---|---|---|
| v1 | `docx_xml_tool_v1/` | archived | Initial DOCX XML export, structure/text index, and explicit `replacements.json` writeback. |
| v2.4 | `docx_xml_tool_v2_4/` | archived | First Markdown editing/writeback workflow with `document.md` and `content_map.json`. |
| v2.4.1 | `docx_xml_tool_v2_4_1/` | archived | Stabilized Markdown markers, richer audit output, endnotes support, complex DOCX tests, and release packaging. |
| v2.4.2 | `docx_xml_tool_v2_4_2/` | archived stable core | Smart Field Fill on top of the safe Markdown roundtrip core. |

Current recommended version:

```text
../docx_xml_tool_v2_5/
```

Archived versions may lack newer safety checks, field-fill workflows, launchers, GUI support, tests, or packaging updates from the current release.

## Archive rule

Before starting a new version, move or copy the previous current package into this folder using its exact versioned name, for example:

```text
../docx_xml_tool_v2_5/ -> archive/docx_xml_tool_v2_5/
```

Then create the new current package at the repository root with the new version name.

## Privacy note

Do not add real exported workspaces, `original.docx`, private documents, logs, generated `dist/`, or generated `edits/` folders to this archive.
