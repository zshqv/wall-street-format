---
name: wall-street-format
description: Formats financial models in Excel and pitchbooks in PowerPoint according to institutional investment banking (IB) and private equity (PE) formatting standards.
---

# Wall Street Formatting Standards (IB & PE)

Transform raw financial data, valuation models, and slide concepts into institutional investment banking standard deliverables.

## 1. Excel Financial Model Formatting Rules

### Color Coding Standards
* **Blue Text (`#0000FF` / RGB `0, 0, 255`):** Hardcoded inputs, historical numbers, and drivers.
* **Black Text (`#000000` / RGB `0, 0, 0`):** Formulas, calculations, and internal sheet links.
* **Green Text (`#008000` / RGB `0, 128, 0`):** External links across workbooks or data feeds.
* **Red Text (`#FF0000` / RGB `255, 0, 0`):** Warnings, audit alerts, or key error indicators.
* **Maroon / Dark Red (`#800000`):** Cross-sheet references within the same workbook.

### Typography & Gridlines
* **Font:** Arial or Calibri, size 9pt or 10pt across the sheet.
* **Gridlines:** Keep gridlines visible or use explicit thin light gray borders (`#D3D3D3`).
* **Alignment:**
  * Text & Line Items: Left-aligned (indent sub-items by 2 spaces).
  * Numbers & Dates: Right-aligned.
  * Headers: Centered or Left-aligned; never use merged cells (use "Center Across Selection" if needed).

### Number Formatting
* **Currency:** `$#,##0` or `$#,##0.0` (Use `($#,##0)` for negative numbers; no minus signs).
* **Percentages:** `0.0%` or `0.00%`.
* **Multiples:** `0.0x` or `0.00x`.
* **Zero Values:** Format zeroes as dashes (`-`).

### Borders & Line Hierarchy
* **Top Border:** Thin single line above totals/subtotals.
* **Bottom Border:** Double bottom line (Accounting Underline) below summary/total outputs.
* **Header Rows:** Dark navy fill (`#002060`) with white bold text.

---

## 2. PowerPoint Presentation Formatting Rules

### Color Palette
* **Primary:** Classic Wall Street Navy (`#0A2540`).
* **Secondary:** Slate Gray (`#4A5568`) / Off-White (`#F7FAFC`).
* **Accent:** Crimson/Burgundy (`#8B0000`) or Gold/Bronze (`#C5A059`) for callouts.

### Slide Anatomy
1. **Action Title (Header):** Write a complete lead thought (e.g., *"Q3 Revenue Grew 14% YoY Driven by Enterprise Expansion"* instead of *"Q3 Results"*).
2. **Context Tracker Line:** Small gray category header above the main title (e.g., *Project Titan | M&A Overview*).
3. **Structured Content Cards:** Organize information into 2, 3, or 4 visual columns with consistent margins.
4. **Callout / Recommendation Box:** Use a shaded bottom bar or side card for key strategic commentary.

---

## 3. Standard Delivery Workflow
1. Identify whether the user needs an Excel Financial Model layout or a PowerPoint Deck outline.
2. Apply the exact color-coded standards, formula hierarchy, and typography rules.
3. Output clean, ready-to-use code blocks (VBA, Python/openpyxl, or Markdown tables) structured according to Wall Street standards.
