# DOCX XML Tool v2.5（本地 Web GUI + DOCX XML 回写工具）

一个轻量、无第三方依赖的本地 Python 工具，用来把 `.docx` 文件展开成 XML/Markdown 工作区，安全地把文本修改回写成新的 Word 文档，并提供本地浏览器 GUI。

[English README](README.md)

## v2.5 新增内容

v2.5 保留 v2.4.2 的 XML 原位保真核心，新增本地浏览器工作台。GUI 只监听本机 `127.0.0.1`，文档不会上传云端。

GUI 包装三条稳定流程：

1. 把 `.docx` 导出成工作区。
2. 把编辑后的 `document.md` 回写成新的 `.docx`。
3. 把 `field_values.yaml` 中的字段值填入新的 `.docx`。

## 快速开始

### 本地拖拽 Web GUI

双击：

```text
launchers/04_start_web_gui.bat
```

浏览器会打开：

```text
http://127.0.0.1:8765/
```

把 `.docx` 文件拖进页面中间的大框。工具会自动：

1. 把上传文件保存到 `gui_uploads/`；
2. 在 `docx_xml_outputs/` 下导出工作区；
3. 显示按钮用于打开 `document.md`、打开 `field_values.yaml`、回写 Markdown、填充字段、打开输出目录。

生成的新 Word 在工作区里：

```text
edits/markdown_output_001.docx
edits/field_fill_output_001.docx
```

命令行启动 GUI 仍然可用：

```bash
python docx_workspace_tool.py --web-gui
python docx_workspace_tool.py --web-gui --no-open
```

### 备用命令行导出

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

### 命令行 Markdown 回写

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

### 命令行字段填充

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --fill-fields
```

## 工作流

```text
DOCX 输入
  ↓
导出工作区
  ↓
original.docx + unpacked/ + pretty_xml/
  ↓
structure.md + section_context.md + text_index.json
  ↓
document.md + content_map.json + fields.yaml + field_values.yaml
  ↓
编辑 document.md 或 field_values.yaml
  ↓
通过 CLI、启动器或本地 Web GUI 回写
  ↓
edits/*.docx + 审计文件
```

原始 `.docx` 永远不会被修改。

## Windows 启动器

```text
launchers/01_export_docx_workspace.bat      # 把 .docx 拖到这里，导出工作区
launchers/02_apply_markdown_to_docx.bat     # 把已导出的工作区文件夹拖到这里，回写 document.md
launchers/03_fill_fields_to_docx.bat        # 把已导出的工作区文件夹拖到这里，回写 field_values.yaml
launchers/04_start_web_gui.bat              # 启动本地浏览器 GUI
```

`03_fill_fields_to_docx.bat` 需要拖入已经导出的工作区文件夹，不能直接拖原始 .docx，不能直接拖原始 `.docx`。正确流程是先运行 `01_export_docx_workspace.bat`，编辑 `field_values.yaml`，再把整个工作区文件夹拖到 `03_fill_fields_to_docx.bat`。

## 工作区文件说明

| 文件或目录 | 作用 |
|---|---|
| `original.docx` | 原始 DOCX 的副本。 |
| `unpacked/` | DOCX ZIP 的原始解包内容，用于保真回写。 |
| `pretty_xml/` | 便于阅读的格式化 XML，只用于阅读。 |
| `manifest.json` | 文件清单、大小、内容类型和 XML 状态。 |
| `structure.md` | 人类可读的文档结构。 |
| `section_context.md` | 标题附近上下文，帮助避免改错区域。 |
| `text_index.json` | 机器可读的 XML 文本记录。 |
| `candidate_fields.yaml` | 旧版候选字段列表。 |
| `document.md` | Markdown 编辑层。 |
| `content_map.json` | 回写必须依赖的 block 到 DOCX XML 映射。 |
| `format_profile.yaml` | 格式保留策略摘要。 |
| `fields.yaml` | 自动识别的字段填充候选。 |
| `field_values.yaml` | 用户编辑后供 `--fill-fields` 使用的字段值。 |
| `edits/` | 生成的新 DOCX 和审计文件。 |

## 安全模型

- 保留原 DOCX 包结构，只重写已映射的 XML part。
- 当 Markdown block 标记缺失、重复或未知时拒绝回写。
- 除非映射文本被编辑，否则保留关系、媒体、页眉、页脚、脚注、尾注等包内容。
- 审计文件记录实际修改的 block，便于人工检查。
- 重要文件最终使用前必须用 Word/WPS/LibreOffice 打开做视觉检查。

## v2.5 不包含

- 云端上传、登录、协作或远程存储。
- MCP 服务集成。
- AI 编辑代理。
- 多文档 RAG 或向量搜索。
- 完整 Word 渲染或版式重建。
- 图片、图表、SmartArt、公式、批注、修订、宏、嵌入文件编辑。

## 命名和打包

- 发布目录名：`docx_xml_tool` 或 `docx_xml_tool_v2_5`。
- 发布 zip 名：`dist/docx_xml_tool_v2_5.zip`。
- 不要把生成的 `docx_xml_outputs/`、`*_docx_xml/`、`edits/`、`__pycache__/`、`.pytest_cache/`、日志或私人文档打包。

创建干净发布 zip：

```bash
python scripts/package_release.py
```

## 测试

```bash
python docx_workspace_tool.py --test
python test_docx_workspace_tool.py
```

## 环境要求

- Python 3.9 或更高版本；已用 Python 3.13 测试。
- 不需要安装第三方 Python 包。
- CLI 可在 Windows、macOS、Linux 使用。
- `.bat` 启动器仅适用于 Windows。

## 许可证

MIT
