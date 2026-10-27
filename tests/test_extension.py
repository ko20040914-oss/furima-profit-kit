from datetime import date
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from generate import build_workbook
from extend import extend_workbook
from recalculate import recalculate


class ExtensionTests(unittest.TestCase):
    def test_existing_data_and_new_last_row_are_calculated(self):
        book = build_workbook(rows=20)
        book['販売記録']['B5'] = '既存データ保持'
        book['手数料マスタ']['C5'] = '編集済みメモ'
        extend_workbook(book, 30)
        sales = book['販売記録']
        for col, value in {'A': date(2026, 10, 3), 'B': '拡張後の商品',
                           'C': 'メルカリ', 'D': 100, 'E': 1000,
                           'G': 100, 'H': 20, 'J': '販売済み'}.items():
            sales[f'{col}34'] = value
        self.assertEqual(sales['B5'].value, '既存データ保持')
        self.assertEqual(book['手数料マスタ']['C5'].value, '編集済みメモ')
        self.assertEqual(sales.auto_filter.ref, 'A4:K34')
        self.assertTrue(any('J5:J34' in str(v.sqref) for v in sales.data_validations.dataValidation))
        self.assertEqual(sales['I34'].data_type, 'f')
        self.assertIn('$I$34', book['月次集計']['G5'].value)
        self.assertIn('$L$34', book['在庫一覧']['F34'].value)
        with tempfile.TemporaryDirectory(prefix='furima-extend-') as tmp:
            path = Path(tmp) / 'extended.xlsx'
            book.save(path)
            recalculate(path)
            values = load_workbook(path, data_only=True)
            self.assertEqual(values['販売記録']['I34'].value, 680)
            self.assertEqual(values['月次集計']['B18'].value, 7500)
            self.assertEqual(values['月次集計']['G18'].value, 2970)
            self.assertEqual(values['在庫一覧']['B5'].value, '商品C')
        with self.assertRaises(ValueError):
            extend_workbook(book, 20)
