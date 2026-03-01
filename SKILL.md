---
name: office-design-toolkit
description: "Bộ kỹ năng Office (DOCX/PPTX/XLSX/PDF) và Design Toolkit tích hợp từ file người dùng cung cấp."
---

# 📂 OFFICE FILE SKILLS & DESIGN DOCUMENT TOOLKIT
## Tổng hợp toàn bộ Skills xử lý file Office + Design chuyên nghiệp

---

# PHẦN 1: DOCX — Word Document Skill
**Path:** `/mnt/skills/public/docx/SKILL.md`

## Overview
File .docx là ZIP archive chứa XML files.

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` hoặc unpack raw XML |
| Create new document | Dùng `docx-js` |
| Edit existing document | Unpack → edit XML → repack |

### Converting .doc to .docx
```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content
```bash
pandoc --track-changes=all document.docx -o output.md
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images
```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes
```bash
python scripts/accept_changes.py input.docx output.docx
```

## Creating New Documents (docx-js)

### Setup
```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        InternalHyperlink, Bookmark, FootnoteReferenceRun, PositionalTab,
        PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
        TabStopType, TabStopPosition, Column, SectionType,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation
```bash
python scripts/office/validate.py doc.docx
```

### Page Size
```javascript
// CRITICAL: docx-js defaults to A4, not US Letter
sections: [{
  properties: {
    page: {
      size: { width: 12240, height: 15840 },  // US Letter in DXA
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    }
  },
  children: [/* content */]
}]
```

**Common page sizes (DXA units, 1440 DXA = 1 inch):**

| Paper | Width | Height | Content Width (1" margins) |
|-------|-------|--------|---------------------------|
| US Letter | 12,240 | 15,840 | 9,360 |
| A4 (default) | 11,906 | 16,838 | 9,026 |

**Landscape:** Pass portrait dimensions, let docx-js swap:
```javascript
size: {
  width: 12240,   // SHORT edge
  height: 15840,  // LONG edge
  orientation: PageOrientation.LANDSCAPE
},
```

### Styles (Override Built-in Headings)
```javascript
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Title")] }),
    ]
  }]
});
```

### Lists (NEVER use unicode bullets)
```javascript
// ❌ WRONG
new Paragraph({ children: [new TextRun("• Item")] })

// ✅ CORRECT
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
    ]
  }]
});
```

### Tables
```javascript
// CRITICAL: Tables need dual widths - columnWidths on table AND width on each cell
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: 4680, type: WidthType.DXA },
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun("Cell")] })]
        })
      ]
    })
  ]
})
```

### Images
```javascript
new Paragraph({
  children: [new ImageRun({
    type: "png",
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" }
  })]
})
```

### Page Breaks
```javascript
new Paragraph({ children: [new PageBreak()] })
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Hyperlinks
```javascript
// External
new Paragraph({
  children: [new ExternalHyperlink({
    children: [new TextRun({ text: "Click here", style: "Hyperlink" })],
    link: "https://example.com",
  })]
})

// Internal (bookmark + reference)
new Paragraph({ heading: HeadingLevel.HEADING_1, children: [
  new Bookmark({ id: "chapter1", children: [new TextRun("Chapter 1")] }),
]})
new Paragraph({ children: [new InternalHyperlink({
  children: [new TextRun({ text: "See Chapter 1", style: "Hyperlink" })],
  anchor: "chapter1",
})]})
```

### Footnotes
```javascript
const doc = new Document({
  footnotes: {
    1: { children: [new Paragraph("Source: Annual Report 2024")] },
  },
  sections: [{
    children: [new Paragraph({
      children: [
        new TextRun("Revenue grew 15%"),
        new FootnoteReferenceRun(1),
      ],
    })]
  }]
});
```

### Tab Stops
```javascript
new Paragraph({
  children: [
    new TextRun("Company Name"),
    new TextRun("\tJanuary 2025"),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
})
```

### Multi-Column Layouts
```javascript
sections: [{
  properties: {
    column: { count: 2, space: 720, equalWidth: true, separate: true },
  },
  children: [/* content flows naturally */]
}]
```

### Table of Contents
```javascript
new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })
```

### Headers/Footers
```javascript
sections: [{
  properties: {
    page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Critical Rules for docx-js
- Set page size explicitly (default A4)
- Landscape: pass portrait dimensions
- Never use `\n` — use separate Paragraph elements
- Never use unicode bullets — use `LevelFormat.BULLET`
- PageBreak must be in Paragraph
- ImageRun requires `type`
- Always set table `width` with DXA (never PERCENTAGE)
- Tables need dual widths: `columnWidths` + cell `width`
- Always add cell margins
- Use `ShadingType.CLEAR` (never SOLID)
- Never use tables as dividers/rules
- TOC requires HeadingLevel only
- Override built-in styles with exact IDs: "Heading1", "Heading2"
- Include `outlineLevel` for TOC
- Do NOT insert internal style/theme guidance text into the document body (e.g., "Theme: ...", "layout guideline ...") unless user explicitly asks to show it as content

## Editing Existing Documents

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```

### Step 2: Edit XML
- Use "Claude" as author for tracked changes
- Use Edit tool directly (not Python scripts)
- Smart quotes: `&#x2019;` (apostrophe), `&#x201C;`/`&#x201D;` (double quotes)

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```

### XML Reference — Tracked Changes
```xml
<!-- Insertion -->
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>

<!-- Deletion -->
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

### Dependencies
- pandoc, docx (`npm install -g docx`), LibreOffice, Poppler

---

# PHẦN 2: PPTX — PowerPoint Skill
**Path:** `/mnt/skills/public/pptx/SKILL.md`

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | Xem editing workflow |
| Create from scratch | Dùng PptxGenJS |

## Reading Content
```bash
python -m markitdown presentation.pptx
python scripts/thumbnail.py presentation.pptx
python scripts/office/unpack.py presentation.pptx unpacked/
```

## Design Ideas — CRITICAL

### Before Starting
- Pick a bold, content-informed color palette
- Dominance over equality (60-70% one color, 1-2 supporting, 1 accent)
- Dark/light contrast: Dark backgrounds for title + conclusion, light for content
- Commit to ONE visual motif and repeat it

### Color Palettes

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| Midnight Executive | `1E2761` | `CADCFC` | `FFFFFF` |
| Forest & Moss | `2C5F2D` | `97BC62` | `F5F5F5` |
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` |
| Ocean Gradient | `065A82` | `1C7293` | `21295C` |
| Charcoal Minimal | `36454F` | `F2F2F2` | `212121` |
| Teal Trust | `028090` | `00A896` | `02C39A` |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` |
| Sage Calm | `84B59F` | `69A297` | `50808E` |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` |

### For Each Slide
- **Every slide needs a visual element** — image, chart, icon, or shape
- Layout options: Two-column, Icon+text rows, 2x2/2x3 grid, Half-bleed image
- Data display: Large stat callouts (60-72pt), Comparison columns, Timeline/process flow

### Typography

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Palatino | Garamond |

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

### Avoid (Common Mistakes)
- Don't repeat same layout
- Don't center body text (left-align paragraphs)
- Don't skimp on size contrast
- Don't default to blue
- Don't create text-only slides
- **NEVER use accent lines under titles** (AI hallmark)

## PptxGenJS — Creating from Scratch

### Setup
```javascript
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

let slide = pres.addSlide();
slide.addText("Hello World!", { x: 0.5, y: 0.5, fontSize: 36, color: "363636" });
pres.writeFile({ fileName: "Presentation.pptx" });
```

### Layout Dimensions
- `LAYOUT_16x9`: 10" × 5.625" (default)
- `LAYOUT_16x10`: 10" × 6.25"
- `LAYOUT_4x3`: 10" × 7.5"
- `LAYOUT_WIDE`: 13.3" × 7.5"

### Text & Formatting
```javascript
slide.addText("Simple Text", {
  x: 1, y: 1, w: 8, h: 2, fontSize: 24, fontFace: "Arial",
  color: "363636", bold: true, align: "center", valign: "middle"
});

// Rich text arrays
slide.addText([
  { text: "Bold ", options: { bold: true } },
  { text: "Italic ", options: { italic: true } }
], { x: 1, y: 3, w: 8, h: 1 });

// Multi-line (requires breakLine: true)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }
], { x: 0.5, y: 0.5, w: 8, h: 2 });
```

### Lists & Bullets
```javascript
// ✅ CORRECT
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ NEVER use unicode bullets
```

### Shapes
```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" }, line: { color: "000000", width: 2 }
});

// Rounded rectangle
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// With shadow
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

### Images
```javascript
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// With options
slide.addImage({
  path: "image.png", x: 1, y: 1, w: 5, h: 3,
  rounding: true, transparency: 50,
  sizing: { type: 'cover', w: 4, h: 3 }
});
```

### Icons (react-icons → PNG)
```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
```

### Charts
```javascript
slide.addChart(pres.charts.BAR, [{
  name: "Sales", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4500, 5500, 6200, 7100]
}], {
  x: 0.5, y: 0.6, w: 6, h: 3, barDir: 'col',
  showTitle: true, title: 'Quarterly Sales',
  chartColors: ["0D9488", "14B8A6", "5EEAD4"],
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
});
```

### Slide Backgrounds
```javascript
slide.background = { color: "F1F1F1" };
slide.background = { path: "https://example.com/bg.jpg" };
slide.background = { data: "image/png;base64,..." };
```

### Common Pitfalls
1. **NEVER use "#" with hex colors** — causes corruption
2. **NEVER encode opacity in hex string** — use `opacity` property
3. **Use `bullet: true`** — NEVER unicode "•"
4. **Use `breakLine: true`** between items
5. **NEVER reuse option objects** (PptxGenJS mutates them)
6. **Don't use ROUNDED_RECTANGLE with accent borders**

## Editing Existing Presentations

### Workflow
1. Analyze: `python scripts/thumbnail.py template.pptx`
2. Plan slide mapping (USE VARIED LAYOUTS!)
3. Unpack: `python scripts/office/unpack.py template.pptx unpacked/`
4. Build: Delete/duplicate/reorder slides
5. Edit content in XML
6. Clean: `python scripts/clean.py unpacked/`
7. Pack: `python scripts/office/pack.py unpacked/ output.pptx --original template.pptx`

### Scripts
| Script | Purpose |
|--------|---------|
| `unpack.py` | Extract and pretty-print |
| `add_slide.py` | Duplicate slide or create from layout |
| `clean.py` | Remove orphaned files |
| `pack.py` | Repack with validation |
| `thumbnail.py` | Visual grid of slides |

## QA (Required!)
```bash
# Content QA
python -m markitdown output.pptx

# Visual QA — convert to images
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

### Dependencies
- `pip install "markitdown[pptx]"` + `Pillow`
- `npm install -g pptxgenjs`
- LibreOffice, Poppler

---

# PHẦN 3: XLSX — Excel Spreadsheet Skill
**Path:** `/mnt/skills/public/xlsx/SKILL.md`

## Requirements for All Excel Files
- Professional font (Arial, Times New Roman)
- ZERO formula errors
- Preserve existing templates

## Financial Models — Color Coding
- **Blue text** (0,0,255): Hardcoded inputs
- **Black text** (0,0,0): ALL formulas
- **Green text** (0,128,0): Links from other worksheets
- **Red text** (255,0,0): External links
- **Yellow background** (255,255,0): Key assumptions

## Number Formatting
- Years: Text strings ("2024" not "2,024")
- Currency: $#,##0, specify units in headers
- Zeros: "-" format
- Percentages: 0.0%
- Negatives: Parentheses (123) not -123

## CRITICAL: Use Formulas, Not Hardcoded Values

```python
# ❌ WRONG
total = df['Sales'].sum()
sheet['B10'] = total

# ✅ CORRECT
sheet['B10'] = '=SUM(B2:B9)'
sheet['C5'] = '=(C4-C2)/C2'
sheet['D20'] = '=AVERAGE(D2:D19)'
```

## Common Workflow
1. Choose tool: pandas (data) or openpyxl (formulas/formatting)
2. Create/Load
3. Modify
4. Save
5. **Recalculate formulas (MANDATORY)**:
```bash
python scripts/recalc.py output.xlsx
```
6. Verify and fix errors

## Creating New Files
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook()
sheet = wb.active
sheet['A1'] = 'Hello'
sheet['B2'] = '=SUM(A1:A10)'
sheet['A1'].font = Font(bold=True, color='FF0000')
sheet['A1'].fill = PatternFill('solid', start_color='FFFF00')
sheet.column_dimensions['A'].width = 20
wb.save('output.xlsx')
```

## Reading & Analysis
```python
import pandas as pd
df = pd.read_excel('file.xlsx')
all_sheets = pd.read_excel('file.xlsx', sheet_name=None)
```

## Best Practices
- **pandas**: Data analysis, bulk operations
- **openpyxl**: Complex formatting, formulas
- Cell indices are 1-based
- `data_only=True` reads values (WARNING: saving loses formulas)
- Always recalculate with `scripts/recalc.py`

---

# PHẦN 4: PDF — PDF Processing Skill
**Path:** `/mnt/skills/public/pdf/SKILL.md`

## Quick Reference

| Task | Best Tool |
|------|-----------|
| Merge PDFs | pypdf |
| Split PDFs | pypdf |
| Extract text | pdfplumber |
| Extract tables | pdfplumber |
| Create PDFs | reportlab |
| OCR scanned | pytesseract |
| Fill PDF forms | pdf-lib or pypdf (FORMS.md) |

## Reading PDFs
```python
from pypdf import PdfReader
reader = PdfReader("document.pdf")
text = ""
for page in reader.pages:
    text += page.extract_text()
```

## Extract Tables
```python
import pdfplumber
with pdfplumber.open("document.pdf") as pdf:
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                print(row)
```

## Merge PDFs
```python
from pypdf import PdfWriter, PdfReader
writer = PdfWriter()
for pdf_file in ["doc1.pdf", "doc2.pdf"]:
    reader = PdfReader(pdf_file)
    for page in reader.pages:
        writer.add_page(page)
with open("merged.pdf", "wb") as output:
    writer.write(output)
```

## Create PDFs (reportlab)
```python
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

doc = SimpleDocTemplate("report.pdf", pagesize=letter)
styles = getSampleStyleSheet()
story = []
story.append(Paragraph("Report Title", styles['Title']))
story.append(Spacer(1, 12))
story.append(Paragraph("Body text here.", styles['Normal']))
doc.build(story)
```

### Subscripts/Superscripts — NEVER use Unicode
```python
# ✅ CORRECT
chemical = Paragraph("H<sub>2</sub>O", styles['Normal'])
squared = Paragraph("x<super>2</super>", styles['Normal'])
```

## Command-Line Tools
```bash
# Extract text
pdftotext input.pdf output.txt
pdftotext -layout input.pdf output.txt

# Merge
qpdf --empty --pages file1.pdf file2.pdf -- merged.pdf

# Split
qpdf input.pdf --pages . 1-5 -- pages1-5.pdf

# Rotate
qpdf input.pdf output.pdf --rotate=+90:1
```

## OCR Scanned PDFs
```python
import pytesseract
from pdf2image import convert_from_path
images = convert_from_path('scanned.pdf')
text = ""
for i, image in enumerate(images):
    text += pytesseract.image_to_string(image)
```

## Watermark & Password
```python
# Watermark
for page in reader.pages:
    page.merge_page(watermark)
    writer.add_page(page)

# Encrypt
writer.encrypt("userpassword", "ownerpassword")
```

---

# PHẦN 5: FRONTEND DESIGN Skill
**Path:** `/mnt/skills/public/frontend-design/SKILL.md`

## Design Thinking Process
1. **Purpose**: What problem? Who uses it?
2. **Tone**: Brutally minimal, maximalist, retro-futuristic, luxury, playful, editorial, art deco, soft/pastel, industrial...
3. **Constraints**: Framework, performance, accessibility
4. **Differentiation**: What's UNFORGETTABLE?

## Aesthetics Guidelines

### Typography
- Choose beautiful, unique fonts — AVOID Arial, Inter, Roboto
- Pair distinctive display font with refined body font

### Color & Theme
- Cohesive aesthetic with CSS variables
- Dominant colors with sharp accents
- NEVER use generic purple gradients on white

### Motion
- CSS-only animations for HTML
- Motion library for React
- Staggered reveals > scattered micro-interactions

### Spatial Composition
- Unexpected layouts, asymmetry, overlap, diagonal flow
- Grid-breaking elements
- Generous negative space OR controlled density

### Backgrounds & Visual Details
- Gradient meshes, noise textures, geometric patterns
- Layered transparencies, dramatic shadows
- Custom cursors, grain overlays

### NEVER
- Generic AI aesthetics (Inter, Roboto, Arial)
- Cliched color schemes (purple gradients)
- Predictable layouts
- Same design repeated across projects

---

# PHẦN 6: THEME FACTORY Skill
**Path:** `/mnt/skills/examples/theme-factory/SKILL.md`

## 10 Pre-set Themes

### 1. Ocean Depths — Professional maritime
- Deep Navy `#1a2332` | Teal `#2d8b8b` | Seafoam `#a8dadc` | Cream `#f1faee`
- Headers: DejaVu Sans Bold | Body: DejaVu Sans

### 2. Sunset Boulevard — Warm vibrant
- Burnt Orange `#e76f51` | Coral `#f4a261` | Warm Sand `#e9c46a` | Deep Purple `#264653`
- Headers: DejaVu Serif Bold | Body: DejaVu Sans

### 3. Forest Canopy — Earth tones
- Forest Green `#2d4a2b` | Sage `#7d8471` | Olive `#a4ac86` | Ivory `#faf9f6`
- Headers: FreeSerif Bold | Body: FreeSans

### 4. Modern Minimalist — Grayscale
- Charcoal `#36454f` | Slate Gray `#708090` | Light Gray `#d3d3d3` | White `#ffffff`
- Headers: DejaVu Sans Bold | Body: DejaVu Sans

### 5. Golden Hour — Autumnal
- Mustard Yellow `#f4a900` | Terracotta `#c1666b` | Warm Beige `#d4b896` | Chocolate `#4a403a`
- Headers: FreeSans Bold | Body: FreeSans

### 6. Arctic Frost — Winter crisp
- Ice Blue `#d4e4f7` | Steel Blue `#4a6fa5` | Silver `#c0c0c0` | Crisp White `#fafafa`
- Headers: DejaVu Sans Bold | Body: DejaVu Sans

### 7. Desert Rose — Sophisticated dusty
- Dusty Rose `#d4a5a5` | Clay `#b87d6d` | Sand `#e8d5c4` | Deep Burgundy `#5d2e46`
- Headers: FreeSans Bold | Body: FreeSans

### 8. Tech Innovation — Bold modern
- Electric Blue `#0066ff` | Neon Cyan `#00ffff` | Dark Gray `#1e1e1e` | White `#ffffff`
- Headers: DejaVu Sans Bold | Body: DejaVu Sans

### 9. Botanical Garden — Organic fresh
- Fern Green `#4a7c59` | Marigold `#f9a620` | Terracotta `#b7472a` | Cream `#f5f3ed`
- Headers: DejaVu Serif Bold | Body: DejaVu Sans

### 10. Midnight Galaxy — Cosmic dramatic
- Deep Purple `#2b1e3e` | Cosmic Blue `#4a4e8f` | Lavender `#a490c2` | Silver `#e6e6fa`
- Headers: FreeSans Bold | Body: FreeSans

## Usage
1. Show `theme-showcase.pdf` to user
2. Ask for choice
3. Read theme file from `themes/` directory
4. Apply colors + fonts consistently
5. Can also create custom themes on-the-fly

---

# PHẦN 7: CANVAS DESIGN Skill
**Path:** `/mnt/skills/examples/canvas-design/SKILL.md`

## Two-Step Process
1. **Design Philosophy Creation** (.md file)
2. **Express on Canvas** (.pdf or .png file)

## Design Philosophy Creation

### Steps
1. **Name the movement** (1-2 words): "Brutalist Joy", "Chromatic Silence"
2. **Articulate the philosophy** (4-6 paragraphs): Space, color, scale, rhythm, hierarchy

### Philosophy Examples
- **"Concrete Poetry"**: Monumental form, bold geometry, sculptural typography
- **"Chromatic Language"**: Color as primary information system, geometric precision
- **"Analog Meditation"**: Texture, breathing room, Japanese photobook aesthetic
- **"Organic Systems"**: Rounded forms, natural clustering, modular growth
- **"Geometric Silence"**: Grid-based precision, Swiss formalism meets Brutalist

### Essential Principles
- VISUAL PHILOSOPHY — aesthetic worldview
- MINIMAL TEXT — sparse, essential-only
- SPATIAL EXPRESSION — ideas through form, not paragraphs
- EXPERT CRAFTSMANSHIP — meticulously crafted, painstaking attention

## Canvas Creation
- Museum/magazine quality
- Repeating patterns and perfect shapes
- Scientific diagram aesthetic
- Limited, intentional color palette
- Sophisticated, NEVER cartoony

## Available Fonts (canvas-fonts directory)
ArsenalSC, BigShoulders, Boldonse, BricolageGrotesque, CrimsonPro, DMMono, EricaOne, GeistMono, Gloock, IBMPlexMono, IBMPlexSerif, InstrumentSans, InstrumentSerif, Italiana, JetBrainsMono, Jura, LibreBaskerville, Lora, NationalPark, NothingYouCouldDo, Outfit, PixelifySans, PoiretOne, RedHatMono, Silkscreen, SmoochSans, Tektur, WorkSans, YoungSerif

---

# PHẦN 8: CHEAT SHEET — QUICK WORKFLOW MAP

## File Type → Skill Mapping

| Need | Skill | Install |
|------|-------|---------|
| Create .docx | DOCX | `npm install -g docx` |
| Edit .docx | DOCX | unpack/edit XML/pack |
| Create .pptx | PPTX | `npm install -g pptxgenjs` |
| Edit .pptx | PPTX | unpack/edit XML/pack |
| Create .xlsx | XLSX | openpyxl (Python) |
| Edit .xlsx | XLSX | openpyxl + recalc.py |
| Create .pdf | PDF | reportlab (Python) |
| Edit .pdf | PDF | pypdf / pdfplumber |
| Beautiful web UI | Frontend Design | HTML/CSS/JS/React |
| Apply themes | Theme Factory | 10 preset themes |
| Art/poster/visual | Canvas Design | reportlab + fonts |

## Universal Design Principles
1. **Never default to generic** — pick bold, context-specific aesthetics
2. **Color dominance** — one color 60-70%, accents sharp
3. **Typography matters** — avoid Arial/Inter/Roboto for design work
4. **Every slide/page needs visual elements** — never text-only
5. **QA is mandatory** — always verify output visually
6. **Formulas over hardcodes** — keep spreadsheets dynamic
7. **Validate everything** — use validate.py/recalc.py scripts

---

## DOCX Style Blueprint (learned from approved client docs)
Use this as default style system when creating professional DOCX plans/proposals unless user asks otherwise.

### A) Typography & Hierarchy
- Body font: **Calibri**
- Body size: **11pt** (`size: 22` in docx half-points)
- Heading 1: **18pt**, bold, centered on cover/title block
- Heading 2: **14pt**, bold, section headers
- Notes/disclaimers: **10pt–11pt**, italic, muted color

### B) Color Tokens
- Primary dark (header/table head): `1A2332`
- Secondary accent (section emphasis): `2D8B8B`
- Text main: `1F2937`
- Muted note text: `475569`
- Table border: `CBD5E1`
- Zebra row background: `F4F8FA`

### B.1) Alternate Client Style Pack (from reference docs)
Use when user asks for the same visual language as approved sample files.
- Heading 1 color: `2A7A78`
- Heading 2 color: `8B6B4A`
- Heading 3 color: `1F4D78`
- Body text: `2D2D2D` / `333333`
- Table accents: `76C893`, `FF9E8A`, `5B4FA0` (use sparingly)
- Header text on dark fill: `FFFFFF`
- Keep contrast high and avoid overusing many accent colors in one page

### C) Page Layout
- Paper: US Letter (`12240 x 15840` DXA) unless user requests A4
- Margins: `1080` DXA each side (≈0.75")
- Spacing: keep consistent paragraph rhythm (before/after spacing, avoid dense blocks)
- Do not use raw `\n` for layout; split into separate `Paragraph` elements

### D) Tables (must-follow)
- Always use dual width strategy:
  - Table `width` with `WidthType.DXA`
  - `columnWidths` on table
  - `width` on each `TableCell`
- Add visible but soft borders (`CBD5E1`)
- Header row:
  - dark fill (`1A2332`)
  - white bold text
  - centered alignment
- Body rows:
  - zebra striping (`FFFFFF` / `F4F8FA`)
  - center only numeric/short fields, left-align descriptive text
- Add cell paddings/margins for readability

### E) Notes & Meta Text
- Notes should be visually subordinate (italic + muted color)
- Do NOT include internal style instructions in output content
  (e.g., no lines like "Theme: ..." or "layout guideline ...") unless user explicitly asks

### F) QA Checklist before delivery
- Header levels are visually distinct
- Table hierarchy is clear and readable on desktop/mobile view
- No orphan styling, no random colors/fonts
- No internal guideline text leaked into document body
- Output opens cleanly in Word/Google Docs/Office viewers



## Execution Standard (Sprint Rollout: Active)

Use the following references for every Office task:
- Workflow policy: `references/workflow-policy.md`
- DOCX design system: `references/design-system-docx.md`
- QA checklists: `references/qa-checklists.md`

Mandatory operating mode:
1. Create/confirm content first.
2. Build structure second.
3. Apply design third.
4. Run QA before delivery.

If user asks for speed, still preserve the phase order with reduced depth.

## Sprint Implementation Status
- Sprint 1 (workflow lock-in): **Implemented**
- Sprint 2 (layout/style system): **Implemented**
- Sprint 3 (QA framework): **Implemented**

Definition of done for Office output:
- Content quality passed
- Structure clarity passed
- Design consistency passed
- Technical openability passed

Brand handling:
- When brand colors are known, apply them through the design system before final styling.
- Current configured brand color preference: lime green `#ABDF00` gradient-first aesthetic.
