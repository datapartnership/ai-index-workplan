# Handoff: Global AI Index — Work Plan

## Overview
A multi-page document presenting the Global AI Index project work plan, covering April 2026 – November 2027 across six phases. The document includes a cover page, phase overview, two landscape timeline tables, a decision points page, a roundtable touchpoints page, and a company review windows page.

## About the Design Files
The files in this bundle are **design references created in HTML** — high-fidelity prototypes showing the intended look and content, not production code to copy directly. The task is to recreate these designs in a target environment (e.g., React, a CMS, or a document generation pipeline) using its established patterns and libraries.

## Fidelity
**High-fidelity**: The HTML file is a pixel-perfect, self-contained document with final colors, typography, spacing, and layout. Recreate it faithfully using the target codebase's patterns.

## Pages / Views

### Page 1 — Cover
- **Purpose**: Title page
- **Layout**: Dark navy (`#001E60`) background, A4 portrait (210mm × 297mm). Decorative circular borders (top-left, bottom-left). Pill-bar SVG pattern on right edge.
- **Key elements**:
  - Date line: "April 2026 — November 2027" in teal (`#3EACAD`), Asap Condensed 9pt, uppercase, letter-spacing 0.18em
  - Title: "Global AI Index Work Plan" in white, Asap Condensed 31pt, weight 600
  - Subtitle: light white 10pt, weight 300
  - Footer: logo (bottom-left), "Work Plan" label (bottom-right)

### Page 2 — Overview
- **Purpose**: Project summary with phase cards and commitment codes
- **Layout**: A4 portrait, white background, 14mm side padding
- **Key elements**:
  - Section eyebrow in teal, H1 in navy, divider bar (teal + orange)
  - 3×2 grid of phase cards (navy header, gray body)
  - Legend row with colored dots
  - Callout box with commitment codes (teal left border)

### Pages 3 & 4 — Timeline Tables (Landscape)
- **Purpose**: Month-by-month work plan across 7 workstreams
- **Layout**: A4 landscape (297mm × 210mm), 8-column table
- **Columns**: Month | Governance & Roundtable | Legal & MDLA | Stakeholder Engagement | Adoption Indicators (C2,C4,C6) | Users and Use Harmonization (C3,C7) | Research & Supplemental (C5,C8) | Publication & Comms (C1)
- **Row types**:
  - Phase header rows: navy background, white text
  - Roundtable rows: teal-tinted (`#EFF9F9`), teal left border on month cell
  - Decision Point rows: orange-tinted (`#FFF6EF`), orange left border
  - Roundtable + Decision rows: blue-tinted (`#EBF3FA`), navy left border
  - White rows (forced): explicit `background: white !important`
  - Standard alternating: white / `#F8F9FC`
- **Text styling**:
  - `.wp-milestone`: bold navy (`#001E60`, weight 700) — key deliverables
  - `.wp-teal`: teal (`#3EACAD`) — supplemental/research items
  - Bullet lists use `·` pseudo-element in teal

### Page 5 — Decision Points
- **Purpose**: Four decision points with timing, inputs, and owners
- **Layout**: A4 portrait, standard table (`std-table`)

### Page 6 — AI Roundtable Touchpoints
- **Purpose**: Roundtable meeting schedule
- **Layout**: A4 portrait, standard table

### Page 7 — Company Review & Approval Windows
- **Purpose**: Review window schedule
- **Layout**: A4 portrait, standard table

## Design Tokens

### Colors
| Token | Hex | Usage |
|-------|-----|-------|
| Navy | `#001E60` | Primary brand, headers, milestones |
| Purple | `#310459` | Legal column accent |
| Teal | `#3EACAD` | Roundtable, supplemental items, dividers |
| Orange | `#DF6B00` | Decision points |
| Orange Light | `#F7951D` | Publication column accent |
| Gray | `#414042` | Body text |
| Gray Mid | `#8C8A8B` | Secondary text |
| Gray Light | `#E7E6E6` | Borders |

### Typography
- **Display/Headers**: Asap Condensed (400, 500, 600, 700)
- **Body**: Open Sans (300, 400, 600, 700)
- **Table content**: 6.8pt / Open Sans
- **Column headers**: 7.5pt / Asap Condensed 600

### Spacing
- Page padding: 14mm sides, 8mm top, 14mm bottom
- Landscape page: 11mm sides, 5mm top, 10mm bottom
- Table cell padding: 2mm 2.5mm

## Assets
- `Partnership_logo_text_color_dark_bg.png` — DDP logo for dark backgrounds (cover page)
- `Partnership_logo_text_color_light_bg.png` — DDP logo for light backgrounds (interior headers)
- Fonts loaded from Google Fonts: Asap Condensed, Open Sans

## Files
- `AI Index Work Plan - Standalone.html` — Complete self-contained document (all fonts and assets bundled inline). Open in any browser to view. Print to PDF for distribution (Cmd+P / Ctrl+P).
- `styled_converter.py` — Python script to convert the HTML file to a formatted Word document (.docx)
- `AI_Index_Work_Plan - table.docx` — Reference document showing the exact formatting and styling
- `AI Index Work Plan - Standalone.docx` — Generated Word document output

## HTML to Word Conversion

### Requirements
```bash
pip install selenium beautifulsoup4 lxml python-docx
```

You'll also need ChromeDriver installed for Selenium to render the JavaScript-heavy HTML.

### Usage
```bash
python3 styled_converter.py
```

This will:
1. Render the HTML file using Chrome/Selenium (with 12-second wait for JavaScript)
2. Extract all content with proper structure
3. Generate a Word document with exact formatting:
   - 2 landscape sections for timeline tables
   - 5 portrait sections for other content
   - All colors, fonts, and styling from the reference document
   - Column-specific header colors
   - Row background colors (white, #F8F9FC, #FFF6EF, #EFF9F9, #EBF3FA)
   - Bold indigo (#001E60) text for:
     - Prefix patterns (TC:, US:, DG:, etc.)
     - wp-milestone spans (milestone events)
   - Teal (#3EACAD) bullets and wp-teal text
   - Body gray (#414042) text
   - Single spacing in timeline cells
   - Merged phase header rows

### Output
The script generates `AI Index Work Plan - Standalone.docx` with complete formatting that matches the reference document.
