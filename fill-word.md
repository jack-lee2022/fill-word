# /fill-word — 将来源文件内容填入 Word 制式模板

## 用法
```
/fill-word                              # 互动模式：逐步提问
/fill-word <模板路径> <来源文件1> [来源文件2 ...]
/fill-word --template <路径> --source <路径> [--source <路径>] [--output <路径>]
```

**用户的输入参数**：$ARGUMENTS

---

## 执行流程

### Step 0 — 解析参数
从 `$ARGUMENTS` 中提取：
- `--template` 或第一个位置参数 → 模板 .docx 路径
- `--source` 或其余位置参数 → 一个或多个来源文件路径（.docx / .pdf / .txt）
- `--output`（可选）→ 输出路径；若未指定，默认为模板文件名加 `_filled` 后缀，保存于同目录

若参数不足，**依序询问**用户：
1. 模板文件路径（.docx）
2. 来源文件路径（可多个，逗号或换行分隔）
3. 输出文件路径（可按 Enter 使用默认值）

---

### Step 1 — 读取来源文件
对每个来源文件，按扩展名选择读取方式：

**`.docx`** → 用 PowerShell COM 提取完整文本：
```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("<路径>")
$text = $doc.Content.Text
$doc.Close()
$word.Quit()
```

**`.pdf`** → 直接用 `Read` 工具读取（支持 PDF 图文提取）

**`.txt` / `.md`** → 直接用 `Read` 工具读取

将所有来源内容整合后，理解其**结构与层次**（项目名称、背景、目标、预算、人员、风险等常见字段）。

---

### Step 2 — 分析模板结构
用以下 Python 片段列出模板所有段落与表格，以便精确定位每个占位符：

```python
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
doc = Document(r"<模板路径>")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"P{i}: {repr(p.text[:120])}")
print("---TABLES---")
for ti, table in enumerate(doc.tables):
    print(f"Table {ti}: {len(table.rows)}×{len(table.columns)}")
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if cell.text.strip():
                print(f"  T{ti}R{ri}C{ci}: {repr(cell.text.strip()[:80])}")
```

识别占位符类型：
- **单行空白**：`___…___`（下划线串）
- **多行空白**：连续多个下划线段落
- **选项框**：`□`（待勾选）
- **表格空白**：空单元格

---

### Step 3 — 内容映射
根据模板各节的**标题/前导文字**，将来源文件内容智能对应到各占位符。

常见映射规则（不限于此，以模板实际结构为准）：
| 模板字段关键词 | 来源内容 |
|---|---|
| 项目名称 | 来源文件标题或项目名称 |
| 立项背景 | 来源的背景/概述章节 |
| 研发目标 / 总体目标 | 来源的目标章节 |
| 阶段性目标 / 实施计划 | 来源的时程/阶段规划 |
| 详细工作内容 | 来源的模块/任务说明 |
| 创新点 | 来源的技术亮点/自动化功能 |
| 技术难点 | 来源的风险管理/难点章节 |
| 预算 / 费用 | 来源的预算章节或费用明细文件 |
| 人员配置 | 来源的人员/团队说明 |
| 验收标准 | 来源的验收/KPI 章节 |
| 风险 | 来源的风险分析章节 |
| 知识产权 / 专利 | 来源中提及的专利/成果 |

对于**无法从来源文件中找到**对应内容的字段（如个人姓名、具体日期、联系方式），填写 `（待定）` 或 `（待确认）`，并在完成后列出清单告知用户。

---

### Step 4 — 生成并执行 Python 填写脚本
将 Step 3 的映射结果写成完整 Python 脚本，保存为临时文件后执行。

脚本必须包含以下辅助函数：

```python
from docx import Document

def set_para(para, new_text):
    """替换段落所有 run 为 new_text，保留首 run 格式。"""
    if not para.runs:
        para.add_run(new_text)
        return
    for i, run in enumerate(para.runs):
        run.text = new_text if i == 0 else ''

def fill_cell(cell, text):
    """向表格单元格写入文字。"""
    for para in cell.paragraphs:
        set_para(para, '')
    cell.paragraphs[0].add_run(text)

def fill_row(table, row_idx, values):
    """按列依次填写表格行。"""
    row = table.rows[row_idx]
    for ci, val in enumerate(values):
        if ci < len(row.cells):
            fill_cell(row.cells[ci], val)

doc = Document(r"<模板路径>")
paras = doc.paragraphs

# --- 按段落索引填写 ---
set_para(paras[N], "填写内容")

# --- 按表格填写 ---
fill_row(doc.tables[T], R, ["列1", "列2", "列3"])

doc.save(r"<输出路径>")
print("✅ 完成")
```

脚本执行后，检查是否输出 `✅ 完成`；若报错，调试并修正后重新执行。

---

### Step 5 — 报告结果
执行成功后，告知用户：
1. **输出文件路径**
2. **已填写字段摘要**（按章节列出主要内容）
3. **待补充字段清单**（需用户手动填写的项目，如姓名、日期、签字等）

---

## 注意事项
- 若系统未安装 `python-docx`，先执行 `pip install python-docx -q`
- 若来源文件为繁体/简体中文混合，内容直接使用，不做转换
- 若模板含多个相似占位符（如多个 `___` 行），以**段落索引**精确定位，避免错误替换
- 表格行不足时，不自动新增行（避免破坏格式），多余内容追加在最后一行或以备注说明
- 脚本执行完毕后，删除临时脚本文件
