# DOCX XML Tool v2.4 Enhanced Overview

这是在第一版 `docx_xml_tool` 基础上叠加 Markdown 编辑/回写能力的 v2.4 版本。核心原则不变：不重建 Word 文档，不修改原始 DOCX，只基于原 DOCX 包和 XML 映射生成新的 `.docx`。

## 相较第一版的改进

| 方面 | 第一版 | v2.4 |
|---|---|---|
| 基础工作区 | `original.docx`、`unpacked/`、`pretty_xml/`、`manifest.json`、`structure.md`、`section_context.md`、`text_index.json`、`candidate_fields.yaml` | 全部保留 |
| 主要编辑方式 | `replacements.json` 指定 XML 文本节点 | 新增编辑 `document.md` 后通过 `content_map.json` 回写 |
| 新增文件 | 无 | `document.md`、`content_map.json`、`format_profile.yaml` |
| 适合场景 | 精确底层节点替换 | 先用 Markdown 阅读/修改，再安全映射回 Word |
| 格式策略 | 保留原 DOCX 包结构 | 继续保留原 DOCX 包结构，Markdown 只是编辑界面 |

## 推荐流程

```text
DOCX 输入
  ↓
拖到 launchers/01_export_docx_workspace.bat 或运行导出命令
  ↓
生成工作区：original.docx + unpacked/ + pretty_xml/ + 索引文件
  ↓
阅读并编辑 document.md
  ↓
保留 <!--docx:block ...--> 定位标记
  ↓
拖工作区到 launchers/02_apply_markdown_to_docx.bat 或运行 --apply-md
  ↓
生成 edits/markdown_output_001.docx 和审计文件
```

## 一键导出

把 `.docx` 拖到：

```text
launchers/01_export_docx_workspace.bat
```

导出后工作区包含第一版全部文件：

```text
original.docx
unpacked/
pretty_xml/
manifest.json
structure.md
section_context.md
text_index.json
candidate_fields.yaml
```

同时新增 v2.4 文件：

```text
document.md
content_map.json
format_profile.yaml
```

## Markdown 回写

编辑工作区里的 `document.md`。可以改正文、表格单元格、页眉、页脚、脚注文字。

不要删除这种定位标记：

```markdown
<!--docx:block b00001-->
```

编辑完后，把工作区文件夹拖到：

```text
launchers/02_apply_markdown_to_docx.bat
```

新 Word 会生成在：

```text
工作区/edits/markdown_output_001.docx
```

也可以用命令行：

```powershell
python docx_workspace_tool.py --apply-md "C:\path\to\workspace"
```

## 旧式 XML 节点回写仍可用

旧函数 `fill_docx_workspace(workspace, replacements_json, output_docx=None)` 没有删除。原来的 `replacements.json` 节点替换方式仍然可用。

## 完整使用方法

1. 把 `.docx` 拖到 `launchers/01_export_docx_workspace.bat`，或运行：

   ```powershell
   python docx_workspace_tool.py "C:\path\to\input.docx" --workspace-root "C:\path\to\docx_xml_outputs"
   ```

2. 打开生成的工作区，先查看：

   ```text
   structure.md
   section_context.md
   text_index.json
   document.md
   content_map.json
   ```

3. 编辑 `document.md`，但不要删除：

   ```markdown
   <!--docx:block b00001-->
   ```

4. 把整个工作区文件夹拖到 `launchers/02_apply_markdown_to_docx.bat`，或运行：

   ```powershell
   python docx_workspace_tool.py "C:\path\to\workspace" --apply-md
   ```

5. 打开生成的 `工作区/edits/markdown_output_001.docx`，人工检查版式、表格、页眉页脚、脚注、图片和相邻章节是否正常。

## GitHub 发布前检查

发布目录里只应包含源码、README、许可证、测试和启动器。不要发布：

```text
docx_xml_outputs/
*_docx_xml/
edits/
__pycache__/
.pytest_cache/
*.log
真实文档或包含 original.docx 的工作区
```

如果需要演示，请使用虚构内容制作样例文档。

## 隐私和版权注意事项

- 工具本身本地运行，不主动上传文件。
- 导出工作区可能包含完整原文、图片、XML、元数据，应视为敏感内容。
- 不要公开私人、客户、公司、学校或受版权保护的文档，除非有授权。
- 分享前删除姓名、电话、邮箱、身份证号、地址、学号、公司名、签名、图片等个人或敏感信息。
- 不要再分发无权分享的版权模板或文档。
- 审计文件可能包含修改前后的文本，也要检查后再分享。

## 当前不足

- 不是完整 Word 渲染或排版引擎。
- XML 顺序不一定等于 Word 视觉顺序。
- Markdown 不是完整保真 Word 表示，只是编辑表层。
- 复杂功能如批注、修订、内容控件、公式、图表、SmartArt、宏、复杂绘图布局等不保证干净往返。
- 更适合文本级编辑，不适合重新设计版式或复刻模板。

## 命名和打包规则

- 对外发布目录建议命名为 `docx_xml_tool` 或 `docx_xml_tool_v2_4`。
- 避免使用 `enhanced`、`final`、`new`、`test` 这类过程性目录名作为最终发布名。
- 导出工作区命名规则：`YYYY-MM-DD_HHMMSS_<原文件名>`。
- Markdown 回写输出命名规则：`edits/markdown_output_001.docx`、`edits/markdown_output_002.docx`，依次递增。
- XML 节点回写输出命名规则：`edits/output_001.docx`、`edits/output_002.docx`，依次递增。
- `docx_xml_outputs/`、`*_docx_xml/`、`edits/`、`__pycache__/`、日志文件都属于过程性产物，不要随源码一起发布。

## 不变原则

原始 DOCX 不会被修改。增强版只在新工作区和 `edits/` 里写入结果。
