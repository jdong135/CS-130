"""
Generic Integration tests adhering to Specification 1. 
"""

import unittest
import io
import sys
import decimal
from context import sheets

MAX_SHEETS_TEST = 100


class Spec1_Tests(unittest.TestCase):
    """
    Tests adhering to specific examples presented in the Project 1
    specification.
    """

    def test_unique_spreadsheet_names(self):
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        with self.assertRaises(ValueError):
            wb.new_sheet("sheet1")

    def test_spreadsheet_name_case(self):
        wb = sheets.Workbook()
        wb.new_sheet("sHeEt")
        self.assertEqual(wb.spreadsheets['sheet'].name, "sHeEt")

    def test_uniqueness_case_sensitive(self):
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        with self.assertRaises(ValueError):
            wb.new_sheet("Sheet1")

    def test_remove_whitespace_cell1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.set_cell_contents("sheet1", " A1", None)

    def test_remove_whitespace_cell2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.set_cell_contents("sheet1", "A1 ", None)

    def test_new_sheet_names(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        self.assertIn('sheet2', wb.spreadsheets)

    def test_basic_extent(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B3', 'abc')
        self.assertEqual(wb.get_sheet_extent('sheet1'), (2, 3))

    def test_extent_changes(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B3', 'abc')
        wb.set_cell_contents('sheet1', 'A5', 'abc')
        wb.set_cell_contents('sheet1', 'B3', None)
        self.assertEqual(wb.get_sheet_extent('sheet1'), (1, 5))

    def test_single_quote1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "''")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "'")

    def test_single_quote2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "''''")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "'''")

    def test_contents_remove_whitespace(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "    a  b c  ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "a  b c")

    def test_zero_len_contents_and_value(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_empty_contents_and_value(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "     ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_empty_cell_contents_and_value(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), None)

    def test_string_leading_whitespace(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "'   a ")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), '   a')
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), '\'   a')

    def test_literal_int(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "2")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(2))

    def test_trailing_zeros(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "2.0500")
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A1'), decimal.Decimal('2.05'))

    def test_string_arithmetic(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "=2 * \"abc\"")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.CellError))
        self.assertTrue(value.get_type() == sheets.CellErrorType.TYPE_ERROR)

    def test_sheetname_access1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet2')
        wb.set_cell_contents('sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "=sheEt2!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_quotes(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet2')
        wb.set_cell_contents('sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "='sheEt2'!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_underscore_noquotes(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('sheet_2')
        wb.set_cell_contents('sheet_2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "=sheEt_2!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_symbols(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('*(sheet2')
        wb.set_cell_contents('*(sheet2', 'A1', '3')
        wb.set_cell_contents('sheet1', 'A1', "='*(sheEt2'!A1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))

    def test_sheetname_access_symbols_no_quotes(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('*(sheet2')
        wb.set_cell_contents('sheet1', 'A1', '=*(sheEt2!A1')
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.CellError))
        self.assertTrue(value.get_type() == sheets.CellErrorType.PARSE_ERROR)

    def test_parse_str_as_num(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '\'123')
        wb.set_cell_contents('sheet1', 'A2', '3')
        wb.set_cell_contents('sheet1', 'A3', '=A1 * A2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A3'), decimal.Decimal(369))

    def test_concat_int_and_str(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '123')
        wb.set_cell_contents('sheet1', 'A2', '\'a')
        wb.set_cell_contents('sheet1', 'A3', '=A1 &  A2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A3'), '123a')

    def test_empty_cell_as_num_is_zero(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '=A1 * 2')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), decimal.Decimal(0))

    def test_empty_cell_as_str_is_empty(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '="abc" & A1')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), "abc")

    def test_unclear_empty_is_zero(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', None)
        wb.set_cell_contents('sheet1', 'A2', '=A1')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A2'), decimal.Decimal(0))

    def test_parse_err_1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=2fva3')
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.CellError))
        self.assertTrue(value.get_type() == sheets.CellErrorType.PARSE_ERROR)

    def test_circular_reference_L(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
                         sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "")
        wb.set_cell_contents("sheet1", "A1", "")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), None)
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), None)
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
                         sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "")
        wb.set_cell_contents("sheet1", "A1", "")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), None)
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), None)

    def test_bad_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=sheet2!A1')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents('sheet1', 'A1', '=ZZZZZ9')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents('sheet1', 'A1', '=Z99999')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.BAD_REFERENCE)

    def test_divide_zero1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.DIVIDE_BY_ZERO)

    def test_divide_zero2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.CellErrorType.DIVIDE_BY_ZERO)

    def test_error_contents(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'E1', '#div/0!')
        wb.set_cell_contents('sheet1', 'E2', '=e1+5')
        value = wb.get_cell_value('sheet1', 'E2')
        self.assertTrue(isinstance(value, sheets.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.CellErrorType.DIVIDE_BY_ZERO)

    def test_error_contents_awr(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'E1', '=#div/0!')
        wb.set_cell_contents('sheet1', 'E2', '=e1+5')
        value = wb.get_cell_value('sheet1', 'E2')
        self.assertTrue(isinstance(value, sheets.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.CellErrorType.DIVIDE_BY_ZERO)

    def test_update_bad_ref_missing_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=Sheet2!A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        wb.new_sheet()
        self.assertEqual(wb.get_cell_value("sheet1", "a1"), decimal.Decimal(0))

    def test_negate_bad_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=Sheet2!A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "=-A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B1").get_type(), sheets.CellErrorType.BAD_REFERENCE)

    def test_delete_middle_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet3", "A1", "=5")
        wb.set_cell_contents("sheet2", "A1", "=sheet3!A1")
        wb.set_cell_contents("sheet1", "A1", "=sheet2!A1")
        wb.del_sheet("sheet2")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet3", "A1"), decimal.Decimal(5))

    def test_negate_circ_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=A2")
        wb.set_cell_contents("sheet1", "A2", "=A3")
        wb.set_cell_contents("sheet1", "A3", "=A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A3").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "=-A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B1").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_negate_div_by_zero(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1/0")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
        wb.set_cell_contents("sheet1", "B1", "=-A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B1").get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

    def test_unary_error_literal(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=-#REF!")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents("sheet1", "A1", "=3 * 5 -#REF!")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents("sheet1", "A2", "=A1 + 5")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2").get_type(), sheets.CellErrorType.BAD_REFERENCE)

    def test_unary_concat_str(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=(3 + 5 * 7 - 37) & \"a\"")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "1a")
        wb.set_cell_contents("sheet1", "A2", "=1 + 1 & \"a\"")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2").get_type(), sheets.CellErrorType.PARSE_ERROR)

    def test_strip_concat_zeros(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1.50 & \"50\"")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "1.505")
        wb.set_cell_contents("sheet1", "A2", "=1.50 & \"050\"")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), "1.5005")
        wb.set_cell_contents("sheet1", "A1", "=\"1.0\" & 0")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "1")

    def test_topo_update_order(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        sys_out = sys.stdout
        new_stdo = io.StringIO()
        sys.stdout = new_stdo
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet1", "A2", "=A1 * 2")
        wb.set_cell_contents("sheet1", "A3", "=A1 * 3")
        wb.set_cell_contents("sheet1", "A4", "=A2 + A3 * 4")
        wb.set_cell_contents("sheet1", "A5", "=A4 * 5")
        wb.set_cell_contents("sheet1", "A6", "=A2 + A5 * 6")
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("sheet1", "A1", "=5")
        output = new_stdo.getvalue()
        sys.stdout = sys_out
        expected = "[('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet1', 'A3'), " \
            "('Sheet1', 'A4'), ('Sheet1', 'A5'), ('Sheet1', 'A6')]\n"
        self.assertEqual(expected, output)

    def test_topo_update_order_multiple_sheets(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        sys_out = sys.stdout
        new_stdo = io.StringIO()
        sys.stdout = new_stdo
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet2", "A2", "=sheet1!A1 * 2")
        wb.set_cell_contents("sheet1", "A3", "=sheet1!A1 * 3")
        wb.set_cell_contents("sheet3", "A4", "=sheet2!A2 + sheet1!A3 * 4")
        wb.set_cell_contents("sheet2", "A5", "=sheet3!A4 * 5")
        wb.set_cell_contents("sheet3", "A6", "=sheet2!A2 + sheet2!A5 * 6")
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("sheet1", "A1", "=5")
        output = new_stdo.getvalue()
        sys.stdout = sys_out
        expected = "[('Sheet1', 'A1'), ('Sheet2', 'A2'), ('Sheet1', 'A3'), " \
            "('Sheet3', 'A4'), ('Sheet2', 'A5'), ('Sheet3', 'A6')]\n"
        self.assertEqual(expected, output)

    def test_string_num_conversion(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'123")
        self.assertIsInstance(wb.get_cell_value("sheet1", "A1"), str)
        wb.set_cell_contents("sheet1", "A2", "123")
        self.assertIsInstance(wb.get_cell_value(
            "sheet1", "A2"), decimal.Decimal)

    def test_string_concat_whitespace(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", " abc ")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "abc")
        wb.set_cell_contents("sheet1", "A2", "=\" abc \" & \" def \"")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), " abc  def ")

    def test_empty_string(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=""')
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "")
        self.assertNotEqual(wb.get_cell_value("sheet1", "A1"), None)

    def test_subtraction_concat_string(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=(1-1) & " is zero"')
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "0 is zero")

    def test_decimal_concat_string(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=1.0 & " is one"')
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "1 is one")

    def test_negative_decimal(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '0.1')
        self.assertEqual(wb.get_cell_value(
            "sheet1", "A1"), decimal.Decimal("-0.1"))


if __name__ == "__main__":
    unittest.main()
