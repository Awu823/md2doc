# Markdown → Corporate Word 转换器

把 Markdown 文档转换为遵循企业模板排版的 Word（.docx）文件。相比直接用
pandoc，本工具**保留模板的封面、目录页、页眉页脚与样式，只替换示例正文**，
并对正文做企业模板风格的后处理（表格、标题编号、列表、正文缩进等）。

---

## 目录结构

```
md2doc/
├─ run.bat                     # 一键构建入口（Windows）
├─ src/
│  └─ md2docx.py               # 通用转换 CLI（可复用，核心工具）
│  └─ desensitize_template.py  # 模板脱敏工具（清除示例内容）
├─ config/
│  └─ md2docx.json             # 当前文档配置
├─ templates/
│  └─ template.docx            # 企业模板（已脱敏）（排版基准）
├─ input/
│  └─ <文档>.md                # 源 Markdown（放入此目录）
├─ output/
│  └─ <文档>.docx              # 生成产物（写入此目录）
├─ docs/
│  └─ README.md                # 本文档
└─ archive/                    # 早期脚本与报告（历史参考，可忽略）
```

---

## 环境需求

| 依赖 | 说明 | 安装 |
| --- | --- | --- |
| Python 3.8+ | 运行工具 | [python.org](https://www.python.org/downloads/) |
| `python-docx` | 读写 .docx | `pip install python-docx` |
| `lxml` | XML 处理 | `pip install lxml` |
| `pandoc` | Markdown → docx | `winget install pandoc` 或官网 |

安装依赖：

```bash
pip install python-docx lxml
```

> `pandoc` 需在 PATH 中；若找不到，脚本会自动查找
> `%LOCALAPPDATA%\Pandoc\pandoc.exe`（`winget` 默认安装位置）。

---

## 快速开始

### 方式 1：一键构建（推荐，Windows）

```powershell
cd C:\path\to\md2doc
.\run.bat
```

- 读取 `config\md2docx.json` → 从 `input\` 读源 MD → 输出到 `output\`。
- 也**可直接双击** `run.bat`（资源管理器）。
- 若在 PowerShell 里提示"不是可识别的命令"，请加 `.\` 前缀。

### 方式 2：显式指定配置文件

```powershell
.\run.bat config\md2docx.json
```

### 方式 3：直接调用 Python CLI（最灵活）

```powershell
python src\md2docx.py --config config\md2docx.json
```

### 成功运行输出示例

```
body start line: 54 | body md lines: 372
pandoc done -> C:\...\_content.docx
saved: C:\...\output\your_document.docx
cleaned temp files
```

---

## 命令行用法

```
python src\md2docx.py [options] [input.md]
```

### 参数一览

| 参数 | 说明 | 默认 |
| --- | --- | --- |
| `input.md`（位置） | 输入 .md 文件路径 | 来自 config |
| `-t, --template FILE` | 企业模板 .docx | `templates\template.docx` |
| `-o, --output FILE` | 输出 .docx 路径 | `output\<输入名>.docx` |
| `--body-start N` | 正文起始行号（0 基） | 自动（第一个 `# ` 标题） |
| `--doc-name TEXT` | 封面 Document Name | 自动从 MD 检测 |
| `--doc-assoc TEXT` | 封面 Associated Documents | 自动从 MD 检测 |
| `--history TEXT` | 封面 History 值 | 自动从 MD 检测 |
| `--cover-name T,R,C` | Document Name 单元格 | `0,0,1` |
| `--cover-assoc T,R,C` | Associated Docs 单元格 | `0,3,1` |
| `--cover-hist T,R,C` | History 单元格 | `1,2,1` |
| `--config FILE` | JSON 配置文件 | — |
| `--keep-temp` | 保留 `_body.md` / `_content.docx` | 清理 |

> `T,R,C` = 表格序号、行、列（均 0 基）。

### 常见场景示例

**① 换一个新文档（用命令行，元数据自动检测）**
```powershell
python src\md2docx.py my_new_doc.md -o output\my_new_doc.docx
```
- 输入默认从 `input\` 找；输出默认进 `output\`。
- Document Name / Associated Documents / History 会**自动**从 MD 封面表提取，
  无需手动填。

**② 指定不同的模板**
```powershell
python src\md2docx.py my_new_doc.md -t templates\another.docx -o output\my_new_doc.docx

python src\md2docx.py "C:...\md2doc\input\doc_md.md" -t templates\template.docx -o output\ds_new.docx
```

**③ 手动指定正文起点（MD 里第一个 `# ` 不是正文开头时）**
```powershell
python src\md2docx.py my_new_doc.md --body-start 40 -o output\my_new_doc.docx
```

**④ 模板封面布局不同时，指定封面单元格坐标**
```powershell
python src\md2docx.py my_new_doc.md --cover-name 0,0,1 --cover-assoc 0,3,1 --cover-hist 1,2,1 -o output\my_new_doc.docx
```

**⑤ 手动覆盖封面元数据（自动检测不准时）**
```powershell
python src\md2docx.py my_new_doc.md --doc-name "我的文档" --doc-assoc "附件A 附件B" --history "1.0 First version" -o output\my_new_doc.docx
```

**⑥ 调试：保留中间产物**
```powershell
python src\md2docx.py my_new_doc.md --keep-temp -o output\my_new_doc.docx
# 生成 _body.md（清理后的正文）和 _content.docx（pandoc 原始输出）
```

---

## 配置文件（config\md2docx.json）

用 JSON 固化全部参数，便于**重复构建**与**版本管理**。

```json
{
  "md": "your_document.md",
  "template": "template.docx",
  "output": "your_document.docx",
  "body_start": null,
  "doc_name": null,
  "doc_assoc": null,
  "history": null,
  "cover_name": [0, 0, 1],
  "cover_assoc": [0, 3, 1],
  "cover_hist": [1, 2, 1]
}
```

### 字段说明

| 字段 | 说明 | 取值 |
| --- | --- | --- |
| `md` | 输入 Markdown 路径 | 相对 `input\` 或绝对路径 |
| `template` | 企业模板 | 相对 `templates\` 或绝对路径 |
| `output` | 输出 docx | 相对 `output\` 或绝对路径 |
| `body_start` | 正文起始行号 | `null`（自动）或整数 |
| `doc_name` | 封面 Document Name | `null`（自动）或字符串 |
| `doc_assoc` | 封面 Associated Documents | `null`（自动）或字符串 |
| `history` | 封面 History 值 | `null`（自动）或字符串 |
| `cover_name` | Document Name 单元格 | `[table, row, col]` |
| `cover_assoc` | Associated Docs 单元格 | `[table, row, col]` |
| `cover_hist` | History 单元格 | `[table, row, col]` |

> **优先级**：命令行参数 > 配置文件 > 内置默认（`null` 表示自动检测）。

**换新文档时**，通常只需改 `md` 和 `output` 两个字段，其余保持自动检测：

```json
{
  "md": "my_new_doc.md",
  "output": "my_new_doc.docx"
}
```
（其余字段缺省即自动/默认值。）

---

## 工作原理（转换流水线）

1. **准备正文**：从 Markdown 提取正文（默认从第一个 `# ` 标题起）。自动修复
   Google Docs 导出的"段落后紧跟 `* ` 列表项缺空行"问题，使列表正确渲染。
2. **转换**：`pandoc -f markdown -t docx --reference-doc=template.docx`。
3. **合并**：复制模板，保留封面（Document Information + History）、目录、页眉
   页脚与样式；把 pandoc 正文注入模板第一个 H1 之后。
4. **后处理**（企业模板风格）：
   - **表格**：6 边单线边框、灰色表头（D9D9D9）水平/垂直居中、紧边距、
     固定布局 + 内容感知列宽；
   - **标题**：Heading 1/2/3 自动编号（1. / 1.1 / 1.1.1），缩进覆盖为 210；
   - **列表**：pandoc 编号重映射到模板编号（无序 `●`/`o`/`▪`，有序 `1.`）；
   - **正文**：顶格、无首行缩进、单倍行距、段后 120；
   - **加粗节标签**（如 `3. Outbound Adapter Class ...`）：段前 180 / 段后 60；
   - **注入缺失样式**：`BlockText`（引用块）、`VerbatimChar`（行内代码等宽）。
5. **收尾**：清理临时文件（除非 `--keep-temp`）。

---

## 输出验证

构建完成后可用以下命令快速检查产物：

```powershell
python -c "import docx; d=docx.Document(r'output\your_document.docx'); print('tables=',len(d.tables),'headings=',sum(1 for p in d.paragraphs if p.style.name.startswith('Heading')),'cover=',repr(d.tables[0].rows[0].cells[1].text.strip()))"
```

- 预期 `tables=13`、`headings=18`、`cover` 为文档名。
- 用 Word / LibreOffice 打开产物，检查封面、目录、页眉页脚、表格样式是否
  与模板一致。

---

## 常见问题

**Q1：`run.bat` 报"不是可识别的命令"？**
A：PowerShell 需要 `.\run.bat`；或在 `cmd` 里直接 `run.bat`；或双击文件。

**Q2：提示 `python` 找不到？**
A：用 `py` 替代（如 `py src\md2docx.py ...`），或把 Python 加入 PATH。

**Q3：提示 `pandoc` 找不到？**
A：脚本会查找 `%LOCALAPPDATA%\Pandoc\pandoc.exe`。若未安装：
`winget install pandoc`。

**Q4：输出里列表变成了字面 `*` 文本？**
A：源 MD 在段落和 `* ` 列表项之间缺空行（Google Docs 导出常见）。工具已
自动修复；若仍出现，手动在 MD 里补空行。

**Q5：封面 Document Name / History 不对？**
A：默认自动从 MD 封面表检测。若模板表格结构与 MD 不一致，用
`--cover-name/assoc/hist`（或配置）指定单元格，或用
`--doc-name/--doc-assoc/--history` 手动覆盖。

**Q6：换了一个模板，正文排版乱？**
A：模板的封面表格结构、标题样式（Heading1/2/3 缩进）、列表编号（numId）
若不同，需调整 `--cover-*` 坐标，或在 `src\md2docx.py` 中适配
`add_heading_numbering` / `fix_lists` / `fix_table`。

---

## 相关文件

- `src\md2docx.py`：核心转换 CLI（可复用，全部逻辑在此）。
- `src\desensitize_template.py`：模板脱敏工具（清除模板内嵌的示例正文/封面值，
  保留封面/History/TOC 结构、标题骨架、页眉页脚、样式与编号；原模板备份到
  `archive\template_original.docx`，已 git 忽略）。
- `config\md2docx.json`：当前文档配置。
- `archive\build_docx.py`：早期单文档版本（仅参考，不再维护）。
