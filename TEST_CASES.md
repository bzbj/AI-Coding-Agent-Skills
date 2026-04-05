# Skills Test Cases

This document provides test cases to verify the skills are working correctly after installation.

---

## Test Case 1: PDF Skill - Read

**Objective**: Extract text from an existing PDF file

**Prerequisites**:
- Have a PDF file available (e.g., a prospectus, annual report, or any PDF)

**Test Steps**:

1. Run the extraction command:
```bash
python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py extract text <your-pdf-file.pdf> -p 1-5
```

2. **Expected Result**:
   - JSON output showing extracted text from pages 1-5
   - No Python import errors
   - Status: "success"

3. **Verify**:
   - Text content is readable
   - Page numbers match requested range
   - No encoding issues with Chinese characters (if applicable)

---

## Test Case 2: PDF Skill - Create

**Objective**: Create a PDF from HTML content

**Test Steps**:

1. Create a simple HTML file:
```bash
cat > /tmp/test.html << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <title>Test PDF</title>
</head>
<body>
    <h1>Hello World</h1>
    <p>This is a test PDF document.</p>
</body>
</html>
EOF
```

2. Convert to PDF:
```bash
python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py html /tmp/test.html /tmp/test-output.pdf
```

3. **Expected Result**:
   - PDF file created at `/tmp/test-output.pdf`
   - File size > 0 bytes
   - Can be opened with PDF reader

---

## Test Case 3: DOCX Skill - Create

**Objective**: Generate a Word document

**Test Steps**:

1. Navigate to skill directory:
```bash
cd ~/.claude/skills/kimi-docx
```

2. Initialize environment (first time only):
```bash
./scripts/docx init
```

3. Edit the Program.cs to create a simple document:
```bash
cat > /tmp/docx-work/Program.cs << 'EOF'
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string outputFile = args.Length > 0 ? args[0] : "/tmp/test-output.docx";

using (WordprocessingDocument doc = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
{
    MainDocumentPart mainPart = doc.AddMainDocumentPart();
    mainPart.Document = new Document();
    Body body = new Body();
    
    body.Append(new Paragraph(new Run(new Text("Hello World from DOCX Skill!"))));
    body.Append(new Paragraph(new Run(new Text("This is a test document."))));
    
    mainPart.Document.Append(body);
    mainPart.Document.Save();
}

Console.WriteLine($"Document created: {outputFile}");
EOF
```

4. Build the document:
```bash
./scripts/docx build /tmp/test-output.docx
```

5. **Expected Result**:
   - Document created at `/tmp/test-output.docx`
   - Build process completes without errors
   - Can be opened in Microsoft Word or LibreOffice

---

## Test Case 4: XLSX Skill - Create

**Objective**: Create and validate an Excel file

**Test Steps**:

1. Create a Python script to generate Excel:
```bash
cat > /tmp/test_excel.py << 'EOF'
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

wb = Workbook()
ws = wb.active
ws.title = "Test Sheet"

# Add headers
ws['A1'] = "Name"
ws['B1'] = "Value"
ws['A1'].font = Font(bold=True)
ws['B1'].font = Font(bold=True)

# Add data
ws['A2'] = "Item 1"
ws['B2'] = 100
ws['A3'] = "Item 2"
ws['B3'] = 200

# Add formula
ws['B4'] = '=SUM(B2:B3)'

wb.save('/tmp/test-output.xlsx')
print("Excel file created: /tmp/test-output.xlsx")
EOF
```

2. Run the script:
```bash
python3 /tmp/test_excel.py
```

3. Validate the Excel file:
```bash
~/.claude/skills/kimi-xlsx/scripts/KimiXlsx validate /tmp/test-output.xlsx
```

4. **Expected Result**:
   - Excel file created at `/tmp/test-output.xlsx`
   - Validation passes (exit code 0)
   - Can be opened in Microsoft Excel or LibreOffice Calc

---

## Test Case 5: XLSX Skill - Inspect

**Objective**: Inspect an existing Excel file structure

**Test Steps**:

1. Use the inspect command on the test file:
```bash
~/.claude/skills/kimi-xlsx/scripts/KimiXlsx inspect /tmp/test-output.xlsx --pretty
```

2. **Expected Result**:
   - JSON output showing sheet names, headers, data ranges
   - No errors
   - Accurate cell references

---

## Test Case 6: Integration Test - Claude Auto-Trigger

**Objective**: Verify Claude automatically uses skills

**Test Steps**:

1. Create a test file:
```bash
echo "Test content" > /tmp/test-trigger.txt
```

2. **Action**: Ask Claude to "analyze the PDF file `/tmp/test-output.pdf`"

3. **Expected Behavior**:
   - Claude should automatically detect `.pdf` extension
   - Claude should use `kimi-pdf` skill to extract content
   - Should NOT attempt to read PDF directly with Read tool

4. **Verify**: Check Claude's response mentions using the skill

---

## Quick Verification Script

Run this script to verify all skills:

```bash
#!/bin/bash

echo "=== AI Coding Agent Skills - Quick Verification ==="
echo ""

# Check 1: PDF Skill
echo "[1/3] Checking PDF Skill..."
if python3 ~/.claude/skills/kimi-pdf/scripts/pdf.py --help >/dev/null 2>&1; then
    echo "✅ PDF skill is accessible"
else
    echo "❌ PDF skill check failed"
fi

# Check 2: DOCX Skill
echo "[2/3] Checking DOCX Skill..."
if [ -f ~/.claude/skills/kimi-docx/scripts/docx ]; then
    echo "✅ DOCX skill is installed"
else
    echo "❌ DOCX skill not found"
fi

# Check 3: XLSX Skill
echo "[3/3] Checking XLSX Skill..."
if [ -f ~/.claude/skills/kimi-xlsx/scripts/KimiXlsx ]; then
    echo "✅ XLSX skill is installed"
else
    echo "❌ XLSX skill not found"
fi

echo ""
echo "=== Verification Complete ==="
echo "Run 'ls -la ~/.claude/skills/' for detailed view"
```

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `ModuleNotFoundError` | Install missing Python packages: `pip3 install openpyxl pandas pikepdf pdfplumber` |
| Permission denied | Make scripts executable: `chmod +x ~/.claude/skills/*/scripts/*` |
| Command not found | Check PATH or use full path to scripts |
| Validation fails | Check SKILL.md for specific requirements |

---

## Success Criteria

✅ All tests passed when:
1. All three skills are detected
2. PDF extraction returns valid text
3. DOCX build completes without errors
4. XLSX validation passes
5. Claude recognizes and uses skills automatically

If any test fails, review the installation steps and dependencies.
