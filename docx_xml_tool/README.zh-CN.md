# DOCX XML Tool（DOCX XML 导出工具）

一个轻量、无第三方依赖的 Python 工具，用来把 `.docx` 文件展开成适合 AI 阅读和后续 XML 原位修改的工作区，并支持把指定文本节点安全回写成新的 Word 文档。

[English README](README.md)

## 为什么需要这个工具

`.docx` 本质上是一个包含 XML 的 ZIP 文件。很多高层 Word 库在保存时可能会重建段落、合并或拆分 run、改变格式甚至影响版式。这个工具采用更保守的方式：

1. 复制原始 DOCX。
2. 解包完整 DOCX ZIP 内容。
3. 生成便于阅读的 pretty XML。
4. 生成结构索引和文本索引，方便 AI 或人工定位。
5. 可选：按指定 XML 文本节点替换内容，生成新的 `.docx`。

原始 `.docx` 不会被修改。

## 功能

- 将 DOCX 导出为包含 `original.docx`、`unpacked/`、`pretty_xml/` 的工作区。
- 生成 `structure.md`，快速查看段落、表格、页眉页脚等结构。
- 生成 `section_context.md`，帮助判断标题附近文本归属，降低改错区域的风险。
- 生成 `text_index.json`，记录表格、行、列、段落、XPath 和文本信息。
- 生成 `candidate_fields.yaml`，识别可能的字段、标签值和占位符。
- 支持按明确 XML 文本节点回写，生成新的 Word 文档。
- 每次回写生成审计文件，记录替换前后内容和 zip 检查结果。
- 支持中文路径、中文文件名和中文文档内容。

## 环境要求

- Python 3.9 或更高版本；已用 Python 3.13 测试。
- 不需要安装第三方 Python 包，只使用 Python 标准库。
- Python 命令行工具可在 Windows、macOS、Linux 上运行。
- `launchers/DOCX_XML_export.bat` 拖拽启动器仅适用于 Windows。
- 如果文件名或文档内容包含中文，建议使用支持 UTF-8 的终端或编辑器。Windows 命令提示符里看到中文乱码时，通常只是控制台显示问题，生成的 `structure.md`、`text_index.json`、`last_run.log` 等文件仍然是 UTF-8；建议用 VS Code 等编辑器打开查看。
- 输入文件必须是真正的 `.docx` 文件；旧版 `.doc` 文件需要先转换为 `.docx`。

## 快速开始

```bash
python docx_xml_exporter.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

输出目录示例：

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

也可以指定固定输出目录：

```bash
python docx_xml_exporter.py "C:/path/to/input.docx" --out "C:/path/to/workspace"
```

使用 `--out` 时不会更新工作区根目录下的 `index.md`。

## Windows 拖拽启动器

仓库内提供 Windows 启动器：

```text
launchers/DOCX_XML_export.bat
```

下载或克隆项目后，可以直接把 `.docx` 文件拖到这个 `.bat` 文件上；也可以双击它，然后粘贴 `.docx` 完整路径。

启动器会把结果写到仓库目录下：

```text
docx_xml_outputs/
```

最近一次运行日志在：

```text
docx_xml_outputs/last_run.log
```

这个启动器避免了 Windows 批处理里常见的中文路径、括号路径问题。

## 工作区文件说明

| 文件或目录 | 作用 |
|---|---|
| `original.docx` | 原始 DOCX 的副本。 |
| `unpacked/` | DOCX ZIP 的原始解包内容，后续回写应基于这里。 |
| `pretty_xml/` | 便于阅读的格式化 XML，只用于阅读，不用于回写。 |
| `manifest.json` | 文件清单、大小、内容类型、XML 格式化状态。 |
| `structure.md` | 人类可读的文档结构：正文段落、表格、页眉、页脚、脚注、尾注。 |
| `section_context.md` | 疑似标题附近的上下文，修改前用于确认文本归属。 |
| `text_index.json` | 机器可读文本索引，包含表格/单元格/段落等定位信息。 |
| `candidate_fields.yaml` | 候选字段和占位符。 |
| `edits/` | 生成的新 DOCX 和编辑审计文件。 |

## 回写文本到 DOCX

`fill_docx_workspace()` 可以替换明确的 XML 文本节点，并生成新的 `.docx`，不会修改原始工作区。

替换说明示例：

```json
{
  "replacements": [
    {
      "label": "姓名",
      "part": "word/document.xml",
      "text_node_indexes": [12],
      "old_text": "姓名：张三",
      "new_text": "姓名：李四"
    }
  ]
}
```

Python 调用示例：

```python
from pathlib import Path
from docx_xml_exporter import fill_docx_workspace

workspace = Path("C:/path/to/workspace")
replacements = workspace / "replacements.json"
output_docx = workspace / "edited.docx"

fill_docx_workspace(workspace, replacements, output_docx)
```

如果 `old_text` 和目标节点实际文本不一致，工具会拒绝写入，避免改错 XML 位置。

## 测试

运行自检：

```bash
python docx_xml_exporter.py --test
```

运行单元测试：

```bash
python test_docx_xml_exporter.py
```

## 安全说明

- 工具不会上传文档。
- 工具不需要联网。
- 原始 DOCX 不会被修改。
- 新生成的 DOCX 会写入 `edits/` 或你指定的输出路径。
- `pretty_xml/` 只用于阅读；需要回写时应使用 `unpacked/`。

## 当前限制

- 这不是完整的 Word 渲染或排版引擎。
- XML 文本顺序不一定等于 Word 视觉顺序，尤其是文本框、绘图层、页眉页脚、重复文本层。
- 当前版本重点是 DOCX XML 导出和明确文本节点替换。
- 还没有实现 Markdown 知识库导出功能。

## 后续路线

- 导出适合知识库读取的 `content.md`。
- 生成 `content_map.json`，把 Markdown 表格/段落映射回 DOCX XML 节点。
- 支持 Markdown 内容同步回 DOCX XML，同时保留原 Word 版式。
- 增加更易用的 fill 子命令。

## 许可证

MIT
