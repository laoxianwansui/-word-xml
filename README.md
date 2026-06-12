# DOCX XML Tool

当前推荐版本：

```text
docx_xml_tool_v2_5/
```

这个仓库保存 DOCX XML Tool 的当前版本、历史版本和后续路线图。工具目标是本地优先地解析 DOCX 底层 XML，导出可读 Markdown，并尽量安全地把文本修改回写成新的 Word 文档。

## 版本选择

| Version | Folder | Status | Purpose |
|---|---|---|---|
| v1 | `archive/docx_xml_tool_v1/` | archived | 初版：DOCX 解包、pretty XML、结构索引、文本索引、`replacements.json` 节点回写。 |
| v2.4 | `archive/docx_xml_tool_v2_4/` | archived | 首个 Markdown 编辑/回写闭环：`document.md` + `content_map.json`。 |
| v2.4.1 | `archive/docx_xml_tool_v2_4_1/` | archived | 更严格的 Markdown 标记校验、更详细审计、尾注支持、复杂 Word 回归测试、发布打包脚本。 |
| v2.4.2 | `archive/docx_xml_tool_v2_4_2/` | archived stable core | Smart Field Fill：`fields.yaml` / `field_values.yaml`、字段填充输出、字段审计、重复风险提示、内容控件高风险识别。 |
| v2.5 | `docx_xml_tool_v2_5/` | current local GUI | 本地 Web GUI，把当前 CLI/拖拽流程包成可视化界面。 |
| v2.6 | planned | planned | MCP 集成，让 Claude 等工具可以读取工作区并调用导出/回写能力。 |
| v2.7 | planned | planned | 单文档本地 AI 编辑，在当前安全回写基础上做辅助修改。 |
| v3.5 | planned | planned | 本地结构化知识库/RAG，基于 `document.md`、`structure.md`、`section_context.md` 和索引文件。 |

## 推荐使用 v2.5

进入当前版本目录：

```bash
cd docx_xml_tool_v2_5
```

启动本地 Web GUI：

```bash
python docx_workspace_tool.py --web-gui
```

Windows 也可以运行：

```text
launchers/04_start_web_gui.bat
```

导出 DOCX 工作区：

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

回写成新的 DOCX：

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

字段填充：

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## Windows 启动器

```text
launchers/01_export_docx_workspace.bat
launchers/02_apply_markdown_to_docx.bat
launchers/03_fill_fields_to_docx.bat
launchers/04_start_web_gui.bat
```

## 发布包和 dist

源码仓库默认不提交生成的 `dist/` 压缩包。需要发布 zip 时，在当前版本目录运行：

```bash
python scripts/package_release.py
```

这会在本地生成：

```text
dist/docx_xml_tool_v2_5.zip
```

## 隐私和版权

不要提交或发布：

- 真实用户文档、简历、合同、论文、报告、证书、截图。
- 导出的工作区，尤其是包含 `original.docx` 的目录。
- `docx_xml_outputs/`、`*_docx_xml/`、`edits/`、`__pycache__/`、`.pytest_cache/`、日志文件。
- 无权分享的版权模板或文档。

如需示例，请使用虚构内容制作样例文档。

## 测试当前版本

```bash
python docx_xml_tool_v2_5/test_docx_workspace_tool.py
python docx_xml_tool_v2_5/docx_workspace_tool.py --test
```

## 后续方向

v2.6 再做 MCP 集成。AI 编辑、多文档 RAG 和知识库能力留到后续版本，不放进 v2.5。
