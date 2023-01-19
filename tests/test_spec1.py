"""
Generic Integration tests adhering to Specification 1. 
"""

import unittest
import os
import sys
import decimal

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import Workbook, lark_module, CellErrorType, CellError


MAX_SHEETS_TEST = 100


class Spec1_Tests(unittest.TestCase):
    def test_unique_spreadsheet_names(self):
        wb = Workbook()
        wb.new_sheet("sheet1")
        with self.assertRaises(ValueError):
            wb.new_sheet("sheet1")

    def test_spreadsheet_name_case(self):
        wb = Workbook()
        wb.new_sheet("sHeEt")
        self.assertEqual(wb.spreadsheets['sheet'].name, "sHeEt")

    def test_uniqueness_case_sensitive(self):
        wb = Workbook()
        wb.new_sheet("sheet1")
        with self.assertRaises(ValueError):
            wb.new_sheet("Sheet1")

    def test_remove_whitespace_cell1(self):
        wb = Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.set_cell_contents("sheet1", " A1", None)

    def test_remove_whitespace_cell2(self):
        wb = Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.set_cell_contents("sheet1", "A1 ", None)

    def test_new_sheet_names(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        self.assertIn('sheet2', wb.spreadsheets)

    def test_max_extent(self):
        wb = Workbook()
        for i in range(MAX_SHEETS_TEST):
            wb.new_sheet()
            wb.set_cell_contents(f"sheet{str(i + 1)}", "ZZZZ9999", "test")
        for i in range(MAX_SHEETS_TEST):
            extent = wb.get_sheet_extent(f"sheet{str(i + 1)}")
            self.assertEqual(extent, (475254, 9999))

    def test_basic_extent(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B3', 'abc')
        self.assertEqual(wb.get_sheet_extent('sheet1'), (2, 3))

    def test_extent_changes(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B3', 'abc')
        wb.set_cell_contents('sheet1', 'A5', 'abc')
        wb.set_cell_contents('sheet1', 'B3', None)
        self.assertEqual(wb.get_sheet_extent('sheet1'), (1, 5))

    def test_single_quote(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "''")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "'")

    def test_single_quote(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "''''")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "'''")

    def test_contents_remove_whitespace(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "    a  b c  ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "a  b c")

    def test_zero_len_contents_and_value(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_empty_contents_and_value(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "     ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_empty_cell_contents_and_value(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_string_leading_whitespace(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "'   a ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), '   a')
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), '\'   a')

    def test_literal_int(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "2")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(2))

    def test_trailing_zeros(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "2.0500")
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A1'), decimal.Decimal('2.05'))

    def test_string_arithmetic(self):
        wb = Workbook()
        wb.new_sheet()
        eval, _ = lark_module.evaluate_expr(
            wb, None, "sheet1", "=2 * \"abc\"")
        self.assertEqual(eval.cell_error.get_detail(), "string arithmetic")

    def test_sheetname_access1(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet2')
        wb.set_cell_contents('sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "=sheEt2!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_quotes(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet2')
        wb.set_cell_contents('sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "='sheEt2'!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_underscore_noquotes(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet_2')
        wb.set_cell_contents('sheet_2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "=sheEt_2!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_symbols(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet('*(sheet2')
        wb.set_cell_contents('*(sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "='*(sheEt2'!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_symbols_no_quotes(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet('*(sheet2')
        eval, _ = lark_module.evaluate_expr(
            wb, None, "sheet1", "=*(sheEt2!A1")
        self.assertEqual(eval.cell_error.get_type(), CellErrorType.PARSE_ERROR)

    def test_parse_str_as_num(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '\'123')
        wb.set_cell_contents('sheet1', 'A2', '3')
        wb.set_cell_contents('sheet1', 'A3', '=A1 * A2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A3'), decimal.Decimal(369))

    def test_concat_int_and_str(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '123')
        wb.set_cell_contents('sheet1', 'A2', '\'a')
        wb.set_cell_contents('sheet1', 'A3', '=A1 &  A2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A3'), '123a')

    def test_empty_cell_as_num_is_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '=A1 * 2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), decimal.Decimal(0))

    def test_empty_cell_as_str_is_empty(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '="abc" & A1')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), "abc")

    def test_unclear_empty_is_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '=A1')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), decimal.Decimal(0))

    def test_parse_err_1(self):
        wb = Workbook()
        wb.new_sheet()
        eval, _ = lark_module.evaluate_expr(
            wb, None, "sheet1", "=2fva3")
        self.assertEqual(eval.cell_error.get_type(), CellErrorType.PARSE_ERROR)

    def test_circular_reference_L(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "")
        wb.set_cell_contents("sheet1", "A1", "")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), None)
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), None)
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "")
        wb.set_cell_contents("sheet1", "A1", "")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), None)
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), None)

    def test_bad_ref(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=sheet2!A1')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents('sheet1', 'A1', '=ZZZZZ9')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents('sheet1', 'A1', '=Z99999')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.BAD_REFERENCE)

    def test_divide_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

    def test_divide_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

    def test_error_contents(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'E1', '#div/0!')
        wb.set_cell_contents('sheet1', 'E2', '=e1+5')
        value = wb.get_cell_value('sheet1', 'E2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_error_contents_awr(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'E1', '=#div/0!')
        wb.set_cell_contents('sheet1', 'E2', '=e1+5')
        value = wb.get_cell_value('sheet1', 'E2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO


if __name__ == "__main__":
    unittest.main()
