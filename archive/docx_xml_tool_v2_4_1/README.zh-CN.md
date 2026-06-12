# DOCX XML Tool v2.4.1（DOCX XML 导出与 Markdown 回写工具）

一个轻量、无第三方依赖的 Python 工具，用来把 `.docx` 文件展开成适合 AI 阅读和后续 XML 原位修改的工作区，导出可读 Markdown，并支持把指定文本安全回写成新的 Word 文档。

[English README](README.md)

## 为什么需要这个工具

`.docx` 本质上是一个包含 XML 的 ZIP 文件。很多高层 Word 库在保存时可能会重建段落、合并或拆分 run、改变格式甚至影响版式。这个工具采用更保守的方式：

1. 复制原始 DOCX。
2. 解包完整 DOCX ZIP 内容。
3. 生成便于阅读的 pretty XML。
4. 生成结构索引和文本索引，方便 AI 或人工定位。
5. 生成可读的 `document.md` 和用于回写定位的 `content_map.json`。
6. 可选：按映射的 XML 文本节点替换内容，生成新的 `.docx`。

原始 `.docx` 不会被修改。

## 第一版 vs v2.4

| 方面 | 第一版 `docx_xml_tool` | v2.4.1 |
|---|---|---|
| 核心导出 | `original.docx`、`unpacked/`、`pretty_xml/`、`manifest.json`、`structure.md`、`section_context.md`、`text_index.json`、`candidate_fields.yaml` | 全部保留 |
| 编辑方式 | 通过 `replacements.json` 明确指定 XML 文本节点替换 | 新增 `document.md` + `content_map.json` 的 Markdown 语义编辑 |
| 格式保留策略 | 基于原 DOCX 包，只修改选定 XML 节点 | 策略不变；Markdown 只是编辑表层，不重建 Word |
| 新增元数据 | 无 | `content_map.json` 和 `format_profile.yaml` |
| 适合场景 | 精确的底层 XML 节点修改 | 先阅读/修改 Markdown，再映射回 DOCX XML |

## 流程图

```text
DOCX 输入
  ↓
导出工作区
  ↓
original.docx + unpacked/ + pretty_xml/
  ↓
structure.md + section_context.md + text_index.json
  ↓
document.md + content_map.json + format_profile.yaml
  ↓
编辑 document.md，不删除 <!--docx:block ...--> 定位标记
  ↓
把 Markdown 回写到原 DOCX 包
  ↓
edits/markdown_output_001.docx + edits/markdown_apply_audit_001.md
```

## 完整使用流程

### 第 1 步：把 DOCX 导出成工作区

当你有一个 `.docx` 文件，想检查底层 XML 或通过 Markdown 修改内容时，先做这一步。

Windows 拖拽方式：

```text
把 input.docx 拖到 launchers/01_export_docx_workspace.bat
```

命令行方式：

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

这会在 `docx_xml_outputs/` 下创建一个带时间戳的工作区。

### 第 2 步：检查生成的工作区

编辑前建议先看这些文件：

- `structure.md`：快速查看文档段落、表格、页眉页脚等结构。
- `section_context.md`：查看标题附近上下文；当 Word 视觉顺序和 XML 顺序可能不一致时尤其重要。
- `text_index.json`：机器可读的 XML 文本定位信息。
- `document.md`：可读的 Markdown 编辑层。
- `content_map.json`：Markdown 回写必须依赖的定位映射文件。

### 第 3 步：编辑 `document.md`

可以在 `document.md` 中修改正文、表格单元格、页眉、页脚、脚注文字。

必须保留这种定位标记：

```markdown
<!--docx:block b00001-->
```

不要重命名或删除 `content_map.json`，它负责把 Markdown 内容映射回 DOCX XML。

### 第 4 步：把 Markdown 回写成新的 DOCX

编辑完 `document.md` 后做这一步。

Windows 拖拽方式：

```text
把导出的整个工作区文件夹拖到 launchers/02_apply_markdown_to_docx.bat
```

命令行方式：

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

默认输出为：

```text
workspace/edits/markdown_output_001.docx
workspace/edits/markdown_apply_audit_001.md
```

### 第 5 步：用 Word 检查结果

用 Microsoft Word 或 WPS/LibreOffice 打开生成的 `.docx`，检查：

- 修改内容是否正确。
- 表格显示是否正常。
- 页眉、页脚、脚注、图片、分页是否仍然可接受。
- 相邻章节有没有被误改。

如果结果不对，继续修改 `document.md`，再运行第 4 步生成下一个编号版本。



```text
docx_workspace_tool.py              # 主 Python 命令：导出、回写、自检
docx_markdown/engine.py             # 内部 Markdown 导出/回写引擎
launchers/01_export_docx_workspace.bat
launchers/02_apply_markdown_to_docx.bat
test_docx_workspace_tool.py
```

普通用户只需要运行 `docx_workspace_tool.py` 或编号启动器。`docx_markdown/engine.py` 是主命令调用的内部模块。

## 功能

- 将 DOCX 导出为包含 `original.docx`、`unpacked/`、`pretty_xml/` 的工作区。
- 生成 `structure.md`，快速查看段落、表格、页眉页脚等结构。
- 生成 `section_context.md`，帮助判断标题附近文本归属，降低改错区域的风险。
- 生成 `text_index.json`，记录表格、行、列、段落、XPath 和文本信息。
- 生成 `candidate_fields.yaml`，识别可能的字段、标签值和占位符。
- 生成 `document.md`，作为可读 Markdown 编辑层。
- 生成 `content_map.json`，把 Markdown 块映射回 DOCX XML 位置。
- 生成 `format_profile.yaml`，记录格式保留策略。
- 支持把 Markdown 修改回写为新的 DOCX，尽量保留原 DOCX 包关系和 XML 结构。
- 保留第一版的 `fill_docx_workspace()` 显式 XML 节点替换能力。
- 支持中文路径、中文文件名和中文文档内容。

## 环境要求

- Python 3.9 或更高版本；已用 Python 3.13 测试。
- 不需要安装第三方 Python 包，只使用 Python 标准库。
- Python 命令行工具可在 Windows、macOS、Linux 上运行。
- `launchers/` 中的拖拽启动器仅适用于 Windows。
- 如果文件名或文档内容包含中文，建议使用支持 UTF-8 的终端或编辑器。
- 输入文件必须是真正的 `.docx` 文件；旧版 `.doc` 文件需要先转换为 `.docx`。

## 快速开始

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --workspace-root "C:/path/to/docx_xml_outputs"
```

输出目录示例：

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

也可以指定固定输出目录：

```bash
python docx_workspace_tool.py "C:/path/to/input.docx" --out "C:/path/to/workspace"
```

使用 `--out` 时不会更新工作区根目录下的 `index.md`。

## Windows 拖拽启动器

```text
launchers/01_export_docx_workspace.bat    # 把 .docx 拖到这里，导出工作区
launchers/02_apply_markdown_to_docx.bat      # 把已导出的工作区文件夹拖到这里，回写 document.md
```

`01_export_docx_workspace.bat` 会把结果写到仓库目录下：

```text
docx_xml_outputs/
```

最近一次导出日志在：

```text
docx_xml_outputs/last_run.log
```

`02_apply_markdown_to_docx.bat` 需要拖入一个已经导出的工作区文件夹，里面应包含 `original.docx`、`document.md` 和 `content_map.json`。新 Word 会生成在该工作区的 `edits/` 目录中。

## Markdown 回写

编辑工作区里的 `document.md`。可以改正文、表格单元格、页眉、页脚、脚注文字。

不要删除这种定位标记：

```markdown
<!--docx:block b00001-->
```

编辑后运行：

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md
```

默认输出为：

```text
workspace/edits/markdown_output_001.docx
workspace/edits/markdown_apply_audit_001.md
```

也可以指定输出路径：

```bash
python docx_workspace_tool.py "C:/path/to/workspace" --apply-md --output "C:/path/to/edited.docx"
```

## 显式 XML 节点回写

第一版替换流程仍然可用。`fill_docx_workspace()` 可以替换明确的 XML 文本节点，并生成新的 `.docx`，不会修改原始工作区。

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
from docx_workspace_tool import fill_docx_workspace

workspace = Path("C:/path/to/workspace")
replacements = workspace / "replacements.json"
output_docx = workspace / "edited.docx"

fill_docx_workspace(workspace, replacements, output_docx)
```

如果 `old_text` 和目标节点实际文本不一致，工具会拒绝写入，避免改错 XML 位置。

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
| `document.md` | 可读 Markdown 编辑层。 |
| `content_map.json` | Markdown 块 ID 到 DOCX XML 位置的映射。 |
| `format_profile.yaml` | 格式保留策略和表格合并信息摘要。 |
| `edits/` | 生成的新 DOCX 和编辑审计文件。 |

## 命名和打包规则

- 对外发布的工具目录建议命名为 `docx_xml_tool` 或 `docx_xml_tool_v2_4_1`。
- 不建议在发布目录名里使用 `enhanced`、`final`、`new`、`test` 或纯日期后缀等过程性词。
- 使用 `--workspace-root` 创建的工作区命名为 `YYYY-MM-DD_HHMMSS_<原文件名>`。
- 使用 `--out` 时写入指定目录，不更新 `index.md`。
- Markdown 回写输出命名为 `edits/markdown_output_001.docx`、`edits/markdown_output_002.docx`，依次递增。
- XML 节点替换输出命名为 `edits/output_001.docx`、`edits/output_002.docx`，依次递增。
- 审计文件和生成的 DOCX 放在同一个 `edits/` 目录下。
- 不要把生成的 `docx_xml_outputs/`、`*_docx_xml/`、`edits/`、`__pycache__/` 或日志文件当作源码一起打包。

## GitHub 发布前检查

上传或发布这个目录前，确认包里只包含源码和文档：

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

不要发布这些生成物或隐私文件：

- `docx_xml_outputs/`
- `*_docx_xml/`
- `edits/`
- `__pycache__/`
- `.pytest_cache/`
- `*.log`
- 真实用户文档、简历、合同、论文、报告、证书、截图
- 包含 `original.docx` 副本的导出工作区

如果需要放示例，请使用人工构造的样例文档，里面只放虚构姓名和虚构内容。

## 隐私和版权注意事项

- 工具本身只在本地运行，不会主动上传文档。
- 导出的工作区可能包含完整原文档、XML、文本、元数据和图片，应整体视为敏感内容。
- 不要把私人、保密、客户、学校、公司或受版权保护的文档工作区提交到 GitHub，除非你有明确授权。
- 分享样例前应移除个人信息：姓名、手机号、邮箱、身份证号、地址、学号、公司名、签名和嵌入图片。
- 不要用这个工具再分发你无权分享的版权模板或文档。
- 审计文件可能包含修改前后的文本，分享前也要检查。

## v2.4.1 稳定范围

v2.4.1 是进入 GUI 前的稳定收口版本。它重点增强复杂 Word 文件的安全边界，但不承诺完整理解所有 Word 版式。

v2.4.1 会做：

- 回写前拒绝缺失的 `<!--docx:block ...-->` 标记。
- 拒绝重复 block 标记。
- 拒绝 `content_map.json` 中不存在的 block id。
- 生成更详细的 Markdown 回写审计，记录修改的 block、part、path、修改前文本、修改后文本。
- 测试页眉、页脚、脚注、尾注的导出和回写。
- 记录 `gridSpan` 和 `vMerge` 表格合并元数据。
- 验证图片等未编辑包内容不会丢失。
- 提供排除生成物和隐私文件的发布 zip 脚本。

## 复杂 Word 安全模型

复杂 Word 可能包含绘图层、文本框、重复可见文本、隐藏 XML 顺序、合并表格、页眉、页脚、脚注、尾注、图片、批注和修订。v2.4.1 的策略是保守处理：

1. 保留原 DOCX 包，只重写已映射的 XML part。
2. 把 `document.md` 当作编辑表层，不当作完整 Word 模型。
3. 当 Markdown 标记和 `content_map.json` 不匹配时拒绝回写。
4. 在审计文件中记录实际修改的 block，便于人工检查。
5. 保留图片等未编辑包内容。
6. 最终文件必须人工打开做视觉检查。

## 知识库准备

`document.md` 可以作为知识库输入，但完整知识库不要放在工具源码目录里。推荐位置：

```text
C:/Users/<你>/Documents/docx_knowledge_base/
```

每个导入文档建议至少保留：

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

如果未来还要回写 DOCX，必须保留 `content_map.json`。多文档 RAG、向量索引、数据库检索属于后续版本，不属于 v2.4.1。

## v2.4.1 暂不实现

- GUI。
- MCP 服务集成。
- AI 编辑代理。
- 多文档 RAG 或向量搜索。
- 自动重建所有绘图层/文本框的视觉顺序。
- 编辑图片、图表、SmartArt、公式、批注、修订、内容控件、宏、嵌入文件。
- DOCX、Markdown、LaTeX、HTML 之间的双向同步。

## 测试

运行自检：

```bash
python docx_workspace_tool.py --test
```

运行单元测试：

```bash
python test_docx_workspace_tool.py
```

创建干净发布 zip：

```bash
python scripts/package_release.py
```

## 安全说明

- 工具不会上传文档。
- 工具不需要联网。
- 原始 DOCX 不会被修改。
- 新生成的 DOCX 会写入 `edits/` 或你指定的输出路径。
- `pretty_xml/` 只用于阅读；需要回写时应使用 `unpacked/` 或映射式 Markdown 回写流程。

## 当前限制

- 这不是完整的 Word 渲染或排版引擎。
- XML 文本顺序不一定等于 Word 视觉顺序，尤其是文本框、绘图层、页眉页脚、重复文本层。
- Markdown 是编辑表层，不是完整保真 Word 表示。
- 格式保留依赖于不删除定位标记，并保持 `content_map.json` 可用。
- 复杂 Word 功能不一定能通过 Markdown 干净往返：批注、修订、内容控件、公式、图表、SmartArt、嵌入文件、宏、复杂绘图布局等。
- 这个工具更适合做文本级修改，不适合重新设计版式、重写样式或从零复刻模板。
- 重要文件使用前，一定要打开生成的 DOCX 做人工视觉检查。

## 许可证

MIT
