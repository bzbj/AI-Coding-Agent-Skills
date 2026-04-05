# AI Coding Agent Skills

A collection of Claude Skills for AI-powered document processing and generation.

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

## Installation

1. Clone this repository:
```bash
git clone https://github.com/bzbj/AI-Coding-Agent-Skills.git
```

2. Copy skills to your Claude skills directory:
```bash
mkdir -p ~/.claude/skills
cp -r AI-Coding-Agent-Skills/kimi-pdf ~/.claude/skills/
cp -r AI-Coding-Agent-Skills/kimi-docx ~/.claude/skills/
cp -r AI-Coding-Agent-Skills/kimi-xlsx ~/.claude/skills/
```

## Usage

Once installed, Claude will automatically detect and use these skills when:
1. You mention file paths with matching extensions (.pdf, .docx, .xlsx, etc.)
2. You use trigger keywords related to document processing

## License

MIT License - Feel free to use and modify for your own projects.
