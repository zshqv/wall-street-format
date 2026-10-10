# Wall Street Format – Claude Skill

An institutional-grade **Claude Skill** designed to automatically format financial models (Excel) and pitchbook slides (PowerPoint) according to investment banking and private equity standards.

## Key Features
- **Strict Color-Coding Logic:** Blue inputs, Black formulas, Green external links, Maroon cross-sheet links.
- **IB Accounting Formatting:** Double-line bottom borders, custom currency/multiple masks, and dash-zero formatting.
- **Action-Oriented Slide Layouts:** Formats PowerPoint structure into action titles, context headers, and multi-column comparison cards.
- **Python Automation:** Includes an `openpyxl` script to programmatically apply all formatting standards to any Excel workbook.

## Quickstart / Installation
1. Copy the contents of `SKILL.md`.
2. Add it to your Claude project or local agent skills directory.
3. Call the skill using prompts like:
   - *"Format this 3-statement financial model using Wall Street IB standards."*
   - *"Convert this outline into an investment banking pitchbook slide layout."*

## Python Script Usage
Apply Wall Street formatting to any Excel file:

```bash
pip install openpyxl
python templates/format_model.py input.xlsx --output formatted_output.xlsx
```

The script applies:
- IB color coding (blue inputs, black formulas, green links)
- Navy header rows with white bold text
- Accounting number formats with parenthetical negatives
- Double-line bottom borders on total rows
- Arial 10pt font standardization

## Formatting Standards Summary

### Excel
| Element | Standard |
|---------|----------|
| Font | Arial / Calibri, 9-10pt |
| Inputs | Blue (`#0000FF`) |
| Formulas | Black (`#000000`) |
| External Links | Green (`#008000`) |
| Cross-Sheet Refs | Maroon (`#800000`) |
| Currency | `$#,##0` with `($#,##0)` negatives |
| Multiples | `0.0x` |
| Zeroes | Displayed as `-` |
| Totals | Double bottom border |

### PowerPoint
| Element | Standard |
|---------|----------|
| Primary Color | Navy `#0A2540` |
| Accent | Burgundy `#8B0000` / Gold `#C5A059` |
| Title Style | Action-oriented lead thought |
| Layout | 2-4 column structured content cards |

## License
This project is licensed under the [MIT License](LICENSE).
