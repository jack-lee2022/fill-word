# fill-word

一個 Claude Code 斜線指令（skill），自動將來源文件的內容填入 Word 制式模板（.docx）。

## 功能

- 支援來源格式：`.docx`、`.pdf`、`.txt`、`.md`
- 智能對應模板欄位（項目名稱、背景、目標、預算、人員等）
- 自動識別占位符：下劃線空白、勾選框（□）、表格空白欄位
- 無法對應的欄位標記為 `（待定）` 並列出清單提醒使用者

## 安裝

將 `fill-word.md` 複製到 Claude Code 的自訂指令目錄：

| 平台 | 安裝路徑 |
|------|---------|
| Windows | `C:\Users\<使用者名稱>\.claude\commands\fill-word.md` |
| macOS / Linux | `~/.claude/commands/fill-word.md` |

複製後重新啟動 Claude Code，即可使用 `/fill-word` 指令。

## 前置需求

需安裝 `python-docx`（首次執行時若未安裝，skill 會自動執行安裝）：

```bash
pip install python-docx
```

## 使用方式

```
/fill-word                                         # 互動模式：逐步提問
/fill-word <模板路徑> <來源文件1> [來源文件2 ...]
/fill-word --template <路徑> --source <路徑> [--source <路徑>] [--output <路徑>]
```

### 範例

```
/fill-word template.docx project_plan.pdf budget.docx
/fill-word --template 立項申請書.docx --source 企劃書.docx --output 立項申請書_filled.docx
```

## 輸出

執行完成後會回報：
1. 輸出檔案路徑
2. 已填寫欄位摘要
3. 需手動補充的欄位清單（如姓名、日期、簽字等）
