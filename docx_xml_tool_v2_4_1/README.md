# DOCX XML Tool v2.4.1

A dependency-free Python tool for unpacking `.docx` files into an AI-friendly XML workspace, exporting readable Markdown, and safely writing selected text changes back into a new Word document while preserving the original DOCX package structure.

[中文说明](README.zh-CN.md)

## Why this exists

`.docx` files are ZIP archives containing XML. Editing them through high-level Word libraries can accidentally rebuild paragraphs, lose run-level formatting, or change layout. This tool keeps the original DOCX structure intact:

1. Copy the original file.
2. Unpack the full DOCX archive.
3. Pretty-print XML for reading.
4. Generate structure and text indexes for AI or manual inspection.
5. Export a readable `document.md` plus `content_map.json` for Markdown-based editing.
6. Create a new DOCX by replacing mapped XML text nodes.

The original `.docx` file is never modified.

## v1 vs v2.4

| Area | v1 `docx_xml_tool` | v2.4.1 |
|---|---|---|
| Core export | `original.docx`, `unpacked/`, `pretty_xml/`, `manifest.json`, `structure.md`, `section_context.md`, `text_index.json`, `candidate_fields.yaml` | Same files retained |
| Editing model | Explicit XML text-node replacement through `replacements.json` | Adds semantic Markdown editing through `document.md` plus `content_map.json` |
| Formatting strategy | Preserve the original DOCX package and edit selected XML nodes | Same preservation strategy; Markdown is only an editing surface |
| New metadata | Not available | `content_map.json` and `format_profile.yaml` |
| Best use | Precise low-level XML edits | Review/edit readable Markdown, then map changes back into DOCX XML |

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
document.md + content_map.json + format_profile.yaml
  ↓
Edit document.md without deleting <!--docx:block ...--> markers
  ↓
Apply Markdown back to original DOCX package
  ↓
edits/markdown_output_001.docx + edits/markdown_apply_audit_001.md
```

## Complete usage workflow

### Step 1: Export a DOCX into a workspace

Use this when you have a `.docx` file and want to inspect or edit it through XML/Markdown.

Windows drag-and-drop:

```text
Drag input.docx onto launchers/01_export_docx_workspace.bat
```

Command line:

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

This creates a timestamped workspace under `docx_xml_outputs/`.

### Step 2: Inspect the generated workspace

Open these files before editing:

- `structure.md` — quick document/table overview.
- `section_context.md` — nearby heading context; use this when visual order may differ from XML order.
- `text_index.json` — machine-readable XML text locations.
- `document.md` — readable Markdown editing surface.
- `content_map.json` — required mapping file for Markdown apply.

### Step 3: Edit `document.md`

Edit normal text, table cell text, header/footer text, or footnote text in `document.md`.

Keep every marker like this intact:

```markdown
<!--docx:block b00001-->
```

Do not rename or delete `content_map.json`; it is how Markdown blocks map back to DOCX XML.

### Step 4: Apply Markdown back to a new DOCX

Use this after editing `document.md`.

Windows drag-and-drop:

```text
Drag the exported workspace folder onto launchers/02_apply_markdown_to_docx.bat
```

Command line:

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

The default output is:

```text
workspace/edits/markdown_output_001.docx
workspace/edits/markdown_apply_audit_001.md
```

### Step 5: Check the result in Word

Open the generated `.docx` in Microsoft Word or WPS/LibreOffice and check:

- The changed text is correct.
- Tables still look right.
- Headers, footers, footnotes, images, and page breaks still look acceptable.
- No neighboring section was accidentally changed.

If something is wrong, edit `document.md` again and run step 4 to create the next numbered output.



```text
docx_workspace_tool.py              # main Python command: export, apply, self-test
docx_markdown/engine.py             # internal Markdown export/apply engine
launchers/01_export_docx_workspace.bat
launchers/02_apply_markdown_to_docx.bat
test_docx_workspace_tool.py
```

Ordinary users should run `docx_workspace_tool.py` or the numbered launchers. `docx_markdown/engine.py` is an internal module used by the main command.

## Features

- Export a DOCX into a workspace with `original.docx`, `unpacked/`, and `pretty_xml/`.
- Generate `structure.md` for quick document/table inspection.
- Generate `section_context.md` to reduce mistakes when nearby visual order differs from XML order.
- Generate `text_index.json` with table, row, cell, paragraph, XPath, and text metadata.
- Generate `candidate_fields.yaml` for likely label-value fields and placeholders.
- Generate `document.md` for readable Markdown editing.
- Generate `content_map.json` to map Markdown blocks back to DOCX XML parts.
- Generate `format_profile.yaml` describing the preservation policy.
- Support Markdown apply back into a new DOCX while preserving the original package relationships and XML structure as much as possible.
- Keep the v1 explicit XML-node replacement flow available through `fill_docx_workspace()`.
- Supports Chinese paths, Chinese filenames, and Chinese document text when run with a UTF-8 capable environment.

## Requirements

- Python 3.9 or newer. Python 3.13 is tested.
- No third-party Python packages are required; the tool only uses the Python standard library.
- Works on Windows, macOS, and Linux for the Python CLI.
- The drag-and-drop launchers in `launchers/` are Windows-only.
- For Chinese filenames or document text, use a UTF-8 capable terminal or editor.
- The input file must be a real `.docx` file. Legacy `.doc` files are not supported unless converted to `.docx` first.

## Quick start

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

The output directory will look like:

```text
docx_xml_outputs/
  index.md
  2026-06-10_153000_input/
    original.docx
    unpacked/
    pretty_xml/
    manifest.json
    structure.md
    section_context.md
    text_index.json
    candidate_fields.yaml
    document.md
    content_map.json
    format_profile.yaml
```

You can also choose a fixed output directory:

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --out "C:/path/to/workspace"
```

When `--out` is used, the workspace index is not updated.

## Windows drag-and-drop launchers

```text
launchers/01_export_docx_workspace.bat    # drag a .docx file here to export
launchers/02_apply_markdown_to_docx.bat      # drag an exported workspace folder here to apply document.md
```

`01_export_docx_workspace.bat` writes outputs to:

```text
docx_xml_outputs/
```

and logs the latest export to:

```text
docx_xml_outputs/last_run.log
```

`02_apply_markdown_to_docx.bat` expects an exported workspace folder containing `original.docx`, `document.md`, and `content_map.json`. It writes the new DOCX under the workspace `edits/` folder.

## Markdown apply

Edit the workspace `document.md`. You can change body paragraphs, table cells, headers, footers, and footnote text.

Do not delete block markers like:

```markdown
<!--docx:block b00001-->
```

Apply the edited Markdown with:

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

The default output is:

```text
workspace/edits/markdown_output_001.docx
workspace/edits/markdown_apply_audit_001.md
```

You can choose an explicit output path:

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md --output "C:/path/to/edited.docx"
```

## Explicit XML-node replacement

The v1 replacement flow is still available. `fill_docx_workspace()` can replace explicit XML text nodes and create a new `.docx` without modifying the original workspace.

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
from docx_workspace_tool import fill_docx_workspace

workspace = Path("C:/path/to/workspace")
replacements = workspace / "replacements.json"
output_docx = workspace / "edited.docx"

fill_docx_workspace(workspace, replacements, output_docx)
```

The function refuses to write when `old_text` does not match the target node. This prevents accidental edits to the wrong XML location.

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
| `document.md` | Readable Markdown editing surface. |
| `content_map.json` | Mapping from Markdown block IDs to DOCX XML locations. |
| `format_profile.yaml` | Summary of formatting preservation policy and table merge metadata. |
| `edits/` | Generated DOCX outputs and edit audit files. |

## Naming and packaging rules

- Keep the distributable tool folder named `docx_xml_tool` or `docx_xml_tool_v2_4_1`.
- Avoid temporary package names such as `enhanced`, `final`, `new`, `test`, or date-only suffixes in released folders.
- Workspaces created with `--workspace-root` use `YYYY-MM-DD_HHMMSS_<source-stem>`.
- `--out` writes to the exact folder supplied and does not update `index.md`.
- Markdown apply outputs use `edits/markdown_output_001.docx`, `edits/markdown_output_002.docx`, and so on.
- XML-node replacement outputs use `edits/output_001.docx`, `edits/output_002.docx`, and so on.
- Audit files stay next to generated DOCX files under `edits/`.
- Do not package generated `docx_xml_outputs/`, `*_docx_xml/`, `edits/`, `__pycache__/`, or log files as source code.

## GitHub publishing checklist

Before uploading or publishing this folder, check that the package contains only source files and documentation:

```text
LICENSE
README.md
README.zh-CN.md
README_V2.4.1.md
.gitignore
docx_workspace_tool.py
test_docx_workspace_tool.py
docx_markdown/
launchers/
```

Do not publish generated or private files:

- `docx_xml_outputs/`
- `*_docx_xml/`
- `edits/`
- `__pycache__/`
- `.pytest_cache/`
- `*.log`
- real user documents, resumes, contracts, papers, reports, certificates, or screenshots
- generated workspaces that include copied `original.docx` files

If you want to include examples, create artificial sample documents with fake names and fake content.

## Privacy and copyright notes

- This tool runs locally and does not upload documents by itself.
- Generated workspaces can contain the full original document, extracted XML, text, metadata, and images. Treat the whole workspace as sensitive.
- Do not commit or publish workspaces created from private, confidential, copyrighted, client-owned, school-owned, or employer-owned documents unless you have permission.
- Remove personal data before sharing examples: names, phone numbers, email addresses, ID numbers, addresses, student numbers, company names, signatures, and embedded images.
- Do not use this tool to redistribute copyrighted templates or documents you do not have the right to share.
- Audit files may contain before/after text. Review them before sharing.

## v2.4.1 stability scope

v2.4.1 is a stabilization release before the GUI work. It improves the safety boundary for complex Word files without claiming full Word layout understanding.

What v2.4.1 does:

- rejects missing `<!--docx:block ...-->` markers before writeback;
- rejects duplicate block markers;
- rejects block ids not present in `content_map.json`;
- writes a richer Markdown apply audit with changed block ids, parts, paths, before text, and after text;
- includes headers, footers, footnotes, and endnotes in Markdown export/apply tests;
- records table merge metadata for `gridSpan` and `vMerge`;
- verifies non-edited package parts such as images are preserved;
- provides a release zip script with generated/private files excluded.

## Complex Word safety model

Complex Word files can contain drawing layers, text boxes, repeated visible text, hidden XML order, merged tables, headers, footers, notes, images, comments, and revision data. v2.4.1 handles this by being conservative:

1. Preserve the original DOCX package and only rewrite mapped XML parts.
2. Treat `document.md` as an editing surface, not a full Word model.
3. Refuse writeback when Markdown markers no longer match `content_map.json`.
4. Record changed blocks in the audit file so edits can be inspected.
5. Preserve non-edited package parts such as media files.
6. Require visual inspection of the generated DOCX before final use.

## Knowledge base preparation

`document.md` can be used as knowledge-base input, but the full knowledge base should live outside this source folder. Recommended location:

```text
C:/Users/<you>/Documents/docx_knowledge_base/
```

For each imported document, keep at least:

```text
original.docx
 document.md
content_map.json
structure.md
section_context.md
text_index.json
metadata.yaml
chunks.jsonl
```

Keep `content_map.json` if you may ever write changes back to DOCX. Full multi-document RAG, vector indexes, and database-backed search are planned for a later version, not v2.4.1.

## Not implemented in v2.4.1

- GUI.
- MCP server integration.
- AI editing agent.
- Multi-document RAG or vector search.
- Automatic visual-order reconstruction for every Word drawing/textbox layout.
- Editing images, charts, SmartArt, equations, comments, tracked changes, content controls, macros, or embedded files.
- Bidirectional sync among DOCX, Markdown, LaTeX, and HTML.

## Testing

Run the self-test:

```bash
python docx_workspace_tool.py --test
```

Run the unit tests:

```bash
python test_docx_workspace_tool.py
```

Create a clean release zip:

```bash
python scripts/package_release.py
```

## Safety notes

- The tool does not upload documents.
- The tool does not use network access.
- The source DOCX is not modified.
- Generated DOCX files are written under `edits/` or the output path you provide.
- `pretty_xml/` is for reading only; write back from `unpacked/` or through the mapped Markdown apply flow.

## Limitations

- This is not a full Word rendering engine.
- XML text order may differ from Word visual order, especially with text boxes, drawing layers, headers, footers, and duplicated content.
- Markdown is an editing surface, not a full-fidelity Word representation.
- Formatting preservation depends on keeping the existing block markers and mapping intact.
- Complex Word features may not round-trip cleanly through Markdown: comments, tracked changes, content controls, equations, charts, SmartArt, embedded files, macros, and unusual drawing layouts.
- The tool is best for text-level edits. It is not intended for redesigning layouts, rewriting styles, or recreating a template from scratch.
- Always open the generated DOCX and visually check important pages before using it as a final file.

## License

MIT
