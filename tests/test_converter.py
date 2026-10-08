import sys
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from workbook_to_md import convert, render_sheet, safe_name, text_value
from openpyxl.worksheet.cell_range import CellRange


def cell(row, col, value, fmt='General', kind='s'):
    return SimpleNamespace(row=row, column=col, value=value, number_format=fmt,
                           data_type=kind, coordinate=f'A{row}')


class Sheet:
    title = 'Recipe'
    sheet_state = 'visible'
    _images = []
    _charts = []
    max_row = 4
    max_column = 3

    def __init__(self, rows, merges=()):
        self.rows = rows
        self.merged_cells = SimpleNamespace(ranges=[CellRange(m) for m in merges])

    def iter_rows(self):
        return iter(self.rows)

    def __getitem__(self, coordinate):
        return cell(1, 1, None)


class ConverterTests(unittest.TestCase):
    def test_order_and_empty_positions(self):
        sheet = Sheet([[cell(2, 2, 'Flour'), cell(2, 3, 0)], [cell(4, 1, 'Bake')]])
        text, _ = render_sheet(sheet, None, 'markdown')
        self.assertIn('|  |  |  |\n|  | Flour | 0 |\n|  |  |  |\n| Bake |  |  |', text)

    def test_html_merge(self):
        text, _ = render_sheet(Sheet([[cell(1, 1, 'Title')]], ['A1:C2']), None)
        self.assertIn('rowspan="2" colspan="3"', text)
        self.assertEqual(text.count('<td'), 1)

    def test_escape_and_multiline(self):
        text, _ = render_sheet(Sheet([[cell(1, 1, '<script>|*\nnext')]]), None, 'markdown')
        self.assertIn('&lt;script&gt;&#124;&#42;<br>next', text)

    def test_recipe_formats(self):
        self.assertEqual(text_value(cell(1, 1, 1.5, '# ?/?')), '1 1/2')
        self.assertEqual(text_value(cell(1, 1, 0.25, '# ?/?')), '1/4')
        self.assertEqual(text_value(cell(1, 1, 0.5, '0%')), '50%')
        self.assertEqual(text_value(cell(1, 1, 7, '000')), '007')

    def test_missing_cache_falls_back(self):
        sheet = Sheet([[cell(1, 1, '=1+1', kind='f')]])
        text, warnings = render_sheet(sheet, sheet, formulas='cached')
        self.assertIn('=1+1', text)
        self.assertEqual(len(warnings), 1)

    def test_empty_and_limit(self):
        self.assertIn('Empty worksheet', render_sheet(Sheet([]), None)[0])
        with self.assertRaises(ValueError):
            render_sheet(Sheet([[cell(4, 3, 'x')]]), None, max_cells=2)

    def test_safe_names(self):
        self.assertEqual(safe_name('../../bad'), '_.._bad')
        self.assertEqual(safe_name('CON'), '_CON')

    def test_export_order_report_and_overwrite_protection(self):
        first = Sheet([[cell(1, 1, 'First')]])
        second = Sheet([[cell(1, 1, 'Second')]])
        second.title = 'Other'
        book = SimpleNamespace(worksheets=[first, second], close=lambda: None)
        with tempfile.TemporaryDirectory() as tmp, patch('workbook_to_md.load_workbook', return_value=book):
            report = convert('source.xlsx', tmp)
            self.assertEqual([s['file'] for s in report['sheets']], ['001-Recipe.md', '002-Other.md'])
            self.assertTrue((Path(tmp) / 'conversion-report.json').exists())
            with self.assertRaises(FileExistsError):
                convert('source.xlsx', tmp)
            convert('source.xlsx', tmp, overwrite=True)


if __name__ == '__main__':
    unittest.main()
