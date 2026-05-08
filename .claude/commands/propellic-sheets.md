Create a fully branded Propellic spreadsheet for: $ARGUMENTS

Write a Python script using openpyxl and run it to generate an actual .xlsx file. Save it to the current working directory with a descriptive kebab-case filename. After running, print the full file path.

---

## Propellic Brand Spec — apply exactly

### Colors (use these exact hex values, strip the # for openpyxl PatternFill)
- Midnight: 152534 — header rows, summary rows, reference tab backgrounds
- Pink: E21A6B — accent rows, key metric highlights, tab color for primary sheets
- White: FFFFFF — standard cell background, text on dark backgrounds
- Blush: F7BFD5 — alternating row shading on data rows
- Berry: 8F1F55 — below-benchmark values, negative deltas, reference tab color

### Fonts
- All cells: Calibri (openpyxl default — Montserrat is not an xlsx embed font)
- Header row: bold, white, size 11
- Section label rows: bold, Pink (E21A6B), size 11, ALL CAPS
- Body cells: normal weight, Midnight (152534), size 11
- Summary rows: bold, white on Midnight background, size 11

### Sheet structure rules
1. Always include a **Summary** tab first (Pink tab color: E21A6B)
2. Additional data tabs get Midnight tab color (152534)
3. Reference/lookup tabs get Berry tab color (8F1F55)
4. Row 1 on every sheet: Midnight fill (152534), white bold text, ALL CAPS column headers, freeze this row
5. Alternate data rows: white (FFFFFF) and Blush (F7BFD5)
6. Summary rows (totals, averages): Midnight fill, white bold text
7. Column widths: auto-size to content (min 12, max 40)
8. All number formats: dates as MM/DD/YYYY, currency as $#,##0, percentages as 0.0%
9. Wrap all formulas in IFERROR()
10. Apply conditional formatting on metric columns:
    - Above benchmark / positive delta: Pink fill (E21A6B), white text
    - Below benchmark / negative delta: Berry fill (8F1F55), white text

### Python implementation pattern

```python
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter

# Color constants
MIDNIGHT = "152534"
PINK     = "E21A6B"
WHITE    = "FFFFFF"
BLUSH    = "F7BFD5"
BERRY    = "8F1F55"

def header_style():
    return Font(bold=True, color=WHITE, size=11)

def midnight_fill():
    return PatternFill("solid", fgColor=MIDNIGHT)

def pink_fill():
    return PatternFill("solid", fgColor=PINK)

def blush_fill():
    return PatternFill("solid", fgColor=BLUSH)

def berry_fill():
    return PatternFill("solid", fgColor=BERRY)

def apply_header_row(ws, columns):
    """Write and style Row 1 with Midnight bg, white bold ALL CAPS headers."""
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name.upper())
        cell.font = header_style()
        cell.fill = midnight_fill()
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

def auto_size_columns(ws, min_width=12, max_width=40):
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, min_width), max_width)

def shade_data_rows(ws, start_row, end_row):
    """Alternate White / Blush on data rows."""
    for row_idx in range(start_row, end_row + 1):
        fill = blush_fill() if row_idx % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        for cell in ws[row_idx]:
            if cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF", WHITE, "00" + WHITE):
                cell.fill = fill

def style_summary_row(ws, row_idx):
    for cell in ws[row_idx]:
        cell.font = Font(bold=True, color=WHITE, size=11)
        cell.fill = midnight_fill()
```

Using this pattern, build out all sheets appropriate to the request: $ARGUMENTS

Think carefully about:
- What tabs make sense for this specific spreadsheet
- What columns and sample/formula rows belong on each tab
- Where conditional formatting adds the most value
- What summary calculations belong on the Summary tab

After the script runs successfully, print: "File saved: <full path>"
