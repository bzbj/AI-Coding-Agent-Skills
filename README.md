# AI Coding Agent Skills

A collection of **Claude Code** skills for AI-powered document processing and generation.

> ⚠️ **Important**: These skills are specifically designed for **Claude Code** (Anthropic's CLI tool), not for general use.

## System Requirements

- **Operating System**: macOS or Windows WSL (Windows Subsystem for Linux)
- **Claude Code**: Must be installed and configured
- **Dependencies**: Python 3, Node.js (for PDF generation)

## Skills Overview

### kimi-pdf
Professional PDF solution for:
- Creating PDFs using HTML+Paged.js (academic papers, reports, documents)
- Processing existing PDFs (extract, merge, split, fill forms)
- Supporting KaTeX math formulas, Mermaid diagrams, three-line tables, citations

**Trigger Conditions:**
- File extensions: `.pdf`
- Keywords: "extract text from pdf", "merge pdf", "split pdf", "fill pdf form", "create pdf", "招股说明书", "年报", "财报"

### kimi-docx
Word document generation and editing for:
- Creating professional documents with covers, charts
- Track-changes editing and comments
- Template-based document filling

**Trigger Conditions:**
- File extensions: `.docx`, `.doc`
- Keywords: "create word", "generate docx", "edit word", "track changes", "cover page", "table of contents"

### kimi-xlsx
Excel spreadsheet manipulation for:
- Advanced data analysis and visualization
- Formula deployment and complex formatting
- Pivot table creation and validation
- Financial data processing

**Trigger Conditions:**
- File extensions: `.xlsx`, `.xls`, `.csv`
- Keywords: "excel", "spreadsheet", "pivot table", "数据透视表", "财务分析", "报表"

---

## Installation

### Step 1: Clone Repository

```bash
git clone https://github.com/bzbj/AI-Coding-Agent-Skills.git
cd AI-Coding-Agent-Skills
```

### Step 2: Install Skills

Copy skills to your Claude skills directory:

```bash
mkdir -p ~/.claude/skills
cp -r kimi-pdf ~/.claude/skills/
cp -r kimi-docx ~/.claude/skills/
cp -r kimi-xlsx ~/.claude/skills/
```

### Step 3: Configure Auto-Trigger Rules

Create or edit `~/.claude/CLAUDE.md` and add the following:

```markdown
## 文件处理路由规则

当用户上传或提及以下文件类型时，优先使用对应 skill，而不是直接凭上下文回答：

- `.pdf` → 必须先调用 `kimi-pdf`
- `.docx` / `.doc` → 必须先调用 `kimi-docx`
- `.xlsx` / `.xls` / `.csv` → 必须先调用 `kimi-xlsx`

规则：
1. 先识别文件类型，再选择对应 skill。
2. 若已存在对应 skill，不要跳过 skill 直接回答。
3. 只有在对应 skill 不可用、失败、或用户明确要求不要使用时，才允许回退到普通处理。

---

## Skill 调用路径

| Skill | 脚本路径 |
|-------|----------|
| kimi-pdf | `~/.claude/skills/kimi-pdf/scripts/pdf.py` |
| kimi-docx | `~/.claude/skills/kimi-docx/scripts/docx` |
| kimi-xlsx | `~/.claude/skills/kimi-xlsx/scripts/KimiXlsx` |

---

## 执行流程

当用户提及 PDF/Word/Excel 文件时：

1. **识别文件类型** → 根据扩展名确定使用哪个 skill
2. **读取 SKILL.md** → 先读取对应 skill 的 SKILL.md 了解用法
3. **执行提取** → 使用脚本提取内容
4. **分析数据** → 基于提取的内容进行后续分析

---

## 常用命令示例

```bash
# PDF 文本提取
python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py extract text <file.pdf> -p <pages>

# Word 文档生成
cd ~/.claude/skills/kimi-docx && ./scripts/docx build output.docx

# Excel 验证
~/.claude/skills/kimi-xlsx/scripts/KimiXlsx validate output.xlsx
```
```

---

## Testing Skills (Highly Recommended)

After installation, **strongly recommend** testing the skills to ensure they work properly.

### Test Checklist

| Test Item | Description | Command/Method |
|-----------|-------------|----------------|
| PDF Read | Extract text from a PDF file | `python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py extract text test.pdf -p 1-5` |
| PDF Create | Generate a PDF from HTML | `python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py html input.html output.pdf` |
| Word Create | Generate a Word document | `cd ~/.claude/skills/kimi-docx && ./scripts/docx build test.docx` |
| Excel Create | Create and validate Excel | `~/.claude/skills/kimi-xlsx/scripts/KimiXlsx validate test.xlsx` |

### Quick Test Script

You can run this quick verification:

```bash
# Check if skills are installed
echo "=== Checking Skills Installation ==="
ls -la ~/.claude/skills/

# Check PDF skill
echo "=== PDF Skill ==="
python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py --help 2>/dev/null || echo "PDF skill check failed"

# Check DOCX skill
echo "=== DOCX Skill ==="
cd ~/.claude/skills/kimi-docx && ./scripts/docx env 2>/dev/null || echo "DOCX skill check failed"

# Check XLSX skill
echo "=== XLSX Skill ==="
~/.claude/skills/kimi-xlsx/scripts/KimiXlsx --help 2>/dev/null || echo "XLSX skill check failed"
```

If any test fails, check:
1. Python 3 is installed (`python3 --version`)
2. Node.js is installed (`node --version`)
3. All dependencies are in place (see individual SKILL.md files)

---

## Usage

Once installed and configured:

1. **Claude will automatically detect** file paths with matching extensions
2. **Auto-trigger** when keywords are mentioned
3. Follow the execution flow defined in CLAUDE.md

## License

MIT License - Feel free to use and modify for your own projects.
