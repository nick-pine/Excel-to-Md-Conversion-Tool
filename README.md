# Workbook to Markdown

Creates one UTF-8 `.md` file per Excel worksheet, in workbook order. Requires Python 3.10 or newer. Runs locally; no AI service or account needed.

## Run

Extract this folder, open a terminal in it, and run:

```sh
python -m pip install -r requirements.txt
python workbook_to_md.py "Recipes.xlsx" -o "recipes-md"
```

On macOS/Linux, use `python3` if `python` is unavailable. Output names look like `001-Pancakes.md`. Originals are never modified. Existing outputs are protected; pass `--overwrite` to deliberately replace them.

## Layout modes

Default `--mode html` writes HTML tables inside Markdown files. This retains merged cells with row/column spans. GitHub and Obsidian can render HTML tables, although appearance depends on the viewer.

For ordinary Markdown pipe tables:

```sh
python workbook_to_md.py "Recipes.xlsx" -o "recipes-md" --mode markdown
```

That mode adds Excel column-letter headers, preserves blank cells, and keeps merged-cell text only at the top-left position. Merged ranges are recorded in a comment.

Both modes preserve row/column order, leading and interior blank rows/columns, multiline text, Unicode, zero values, and hidden worksheet/cell data. Trailing empty space and style-only cells are omitted. Ingredient, amount, and unit columns remain separate. Data is escaped so text containing pipes or HTML does not break tables.

## Formulas and quantities

Formula text is preserved by default. Use `--formulas cached` for results last saved by Excel; missing cached results fall back to formulas with a warning. This tool does not calculate formulas. Open, recalculate, and save in Excel before exporting cached values.

Common fraction, fixed-decimal, leading-zero, and percent formats are supported. Dates/times use ISO text. Complex custom Excel number formats are exported as underlying values, so display text may differ. Plain text fractions stay unchanged.

## Preservation limits

This preserves the cell grid and content, not Excel's exact printed appearance. Fonts, colors, borders, widths, row heights, conditional formatting, comments, hyperlinks, drawings, charts, and images are not reproduced. Images/charts on nonempty sheets trigger report warnings. Keep the original workbook alongside your Markdown archive.

Supported inputs: `.xlsx` and `.xlsm` (macros are never executed). Save `.xls` or `.xlsb` as `.xlsx` first. All worksheets, including hidden ones, are exported. Chart-only sheets are not worksheets and are excluded.

`conversion-report.json` maps worksheet names to files and lists warnings. The default grid limit is 250,000 cells per sheet to catch stray formatting and oversized exports; change it with `--max-cells 500000` after inspecting the workbook.

## Verification

```sh
python -m unittest discover -s tests -v
```

Tests use in-memory sheet doubles to verify layout, quantities, escaping, formula fallback, output order, and overwrite protection. No sample recipe workbook was supplied, so validate one representative workbook before converting your whole collection.
