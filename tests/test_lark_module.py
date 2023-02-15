"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
import decimal
from context import sheets


class LarkModuleTests(unittest.TestCase):
    """
    Basic isolated unit tests for lark_module.py
    """

    def test_single_dec(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(wb, None, "sheet1", "=1")
        self.assertEqual(value, 1)

    def test_single_str(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=\"str\"")
        self.assertEqual(value, "str")

    def test_basic1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=1 + 7 - 124 * 7 / 14")
        self.assertEqual(value, -54)

    def test_basic2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=2 * 6 - 4 + 2 / 2")
        self.assertEqual(value, 9)

    def test_basic_parens(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 - 7 + (4 * 8 - 7) - 2 / 2 + (7 - 4)")
        self.assertEqual(value, 24)

    def test_nested_parens(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 / 2 * (7 + 8 - (6 + (2 * 4)) - 4) + 7")
        self.assertEqual(value, 1)

    def test_negative1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1")
        self.assertEqual(value, -1)

    def test_negative2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1 + 2 * (-2 - 2 / (1 * -4))")
        self.assertEqual(value, -4)

    def test_cell_ref1(self):
        wb = sheets.Workbook()
        wb.new_sheet("Sheet1")
        wb.set_cell_contents("Sheet1", "A1", "=4")
        wb.set_cell_contents("Sheet1", "A2", "=A1 + 1")
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "Sheet1", "=A2")
        self.assertEqual(value, 5)

    def test_cell_ref2(self):
        wb = sheets.Workbook()
        wb.new_sheet("Sheet1")
        wb.new_sheet("Sheet2")
        wb.set_cell_contents("Sheet1", "A1", "=1")
        wb.set_cell_contents("Sheet2", "C5", "=3")
        wb.set_cell_contents("Sheet2", "E1", "=Sheet2!C5 + C5 + Sheet1!A1")
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "Sheet2", "=E1")
        self.assertEqual(value, 7)

    def test_zero_div1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "= 2 * 8 / 0 + 1")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    def test_zero_div2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet1", "A2", "=0")
        wb.set_cell_contents("sheet1", "A3", "=A1/A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    def test_string_arithmetic1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'1")
        wb.set_cell_contents("sheet1", "A2", "'hello world")
        wb.set_cell_contents("sheet1", "A3", "=A1 + A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_string_arithmetic2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'goodbye world")
        wb.set_cell_contents("sheet1", "A2", "'hello world")
        wb.set_cell_contents("sheet1", "A3", "=A1 + A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_string_arithmetic3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'10")
        wb.set_cell_contents("sheet1", "A2", "'2")
        wb.set_cell_contents("sheet1", "A3", "=A1 + A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertEqual(value, 12)

    def test_string_arithmetic4(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'1")
        wb.set_cell_contents("sheet1", "A2", "'hello world")
        wb.set_cell_contents("sheet1", "A3", "=A1 * A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_string_arithmetic5(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'goodbye world")
        wb.set_cell_contents("sheet1", "A2", "'hello world")
        wb.set_cell_contents("sheet1", "A3", "=A1 * A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_string_arithmetic6(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'10")
        wb.set_cell_contents("sheet1", "A2", "'2")
        wb.set_cell_contents("sheet1", "A3", "=A1 * A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertEqual(value, 20)

    def test_string_arithmetic7(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'123")
        wb.set_cell_contents("sheet1", "A2", "5.3")
        wb.set_cell_contents("sheet1", "A3", "=A1 * A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertEqual(value, decimal.Decimal('651.9'))

    def test_string_arithmetic8(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'  123")
        wb.set_cell_contents("sheet1", "A2", "5.3")
        wb.set_cell_contents("sheet1", "A3", "=A1 * A2")
        value = wb.get_cell_value('sheet1', 'A3')
        self.assertEqual(value, decimal.Decimal('651.9'))

    def test_bad_reference1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'10")
        wb.set_cell_contents("sheet1", "A2", "=Sheet2!A1")
        value = wb.get_cell_value('sheet1', 'A2')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_bad_reference2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'10")
        wb.set_cell_contents("sheet1", "A2", "=Sheet1!A10000")
        value = wb.get_cell_value('sheet1', 'A2')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_bad_reference3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'10")
        wb.set_cell_contents("sheet1", "A2", "=Sheet1!ZZZZZ9999")
        value = wb.get_cell_value('sheet1', 'A2')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_parse_error1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=e71s18d')
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.PARSE_ERROR)

    def test_error_propagation1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "abc")
        wb.set_cell_contents("sheet1", "B1", "=A1 + 1")
        wb.set_cell_contents("sheet1", "C1", "=B1 + 1")
        value = wb.get_cell_value("sheet1", "C1")
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_error_propagation2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "0")
        wb.set_cell_contents("sheet1", "B1", "=1 / A1")  # Div/0!
        wb.set_cell_contents("sheet1", "C1", "=B1 + 1")
        value = wb.get_cell_value("sheet1", "C1")
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    def test_add(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=1+1")
        self.assertEqual(value, decimal.Decimal(2))

    def test_mult(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=8*2")
        self.assertEqual(value, decimal.Decimal(16))

    def test_neg(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", "=(-5 * 2) * -2")
        self.assertEqual(value, decimal.Decimal(20))

    def test_concat(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", '="abc" & "def" & "ghi"')
        self.assertEqual(value, "abcdefghi")

    def test_parens(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        _, value = sheets.lark_module.evaluate_expr(
            wb, None, "sheet1", '=(1 * 3) + -1 * (3 * -1)')
        self.assertEqual(value, decimal.Decimal(6))

    def test_circ_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        value = wb.get_cell_value('sheet1', 'B1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_three_circ_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=C1')
        wb.set_cell_contents('sheet1', 'C1', '=A1')
        value = wb.get_cell_value('sheet1', 'C1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents('sheet1', 'C1', '1')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal(1))

    def test_circ_ref_update(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        wb.set_cell_contents('sheet1', 'B1', None)
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertEqual(value, decimal.Decimal(0))

    def test_valid_reference_order(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B1', '=C1')
        wb.set_cell_contents('sheet1', 'A1', '=B1 + C1')
        wb.set_cell_contents('sheet1', 'C1', '1')
        values = [wb.get_cell_value('sheet1', 'A1'), wb.get_cell_value(
            'sheet1', 'B1'), wb.get_cell_value('sheet1', 'C1')]
        self.assertEqual(values, [2, 1, 1])

    def test_diamond_dependency(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1 + D1')
        wb.set_cell_contents('sheet1', 'B1', '=C1')
        wb.set_cell_contents('sheet1', 'D1', '=C1')
        wb.set_cell_contents('sheet1', 'C1', '10')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A1'), decimal.Decimal(20))
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'B1'), decimal.Decimal(10))
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'D1'), decimal.Decimal(10))

    def test_first_dollar_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet1", "A2", "=A1")
        wb.set_cell_contents("sheet1", "A3", "=$A1")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"),
                         wb.get_cell_value("sheet1", "A3"))

    def test_second_dollar_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet1", "A2", "=A1")
        wb.set_cell_contents("sheet1", "A3", "=A$1")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"),
                         wb.get_cell_value("sheet1", "A3"))

    def test_double_dollar_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet1", "A2", "=A1")
        wb.set_cell_contents("sheet1", "A3", "=$A$1")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"),
                         wb.get_cell_value("sheet1", "A3"))

    def test_first_dollar_sheet_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet2", "A2", "=sheet1!A1")
        wb.set_cell_contents("sheet2", "A3", "=sheet1!$A1")
        self.assertEqual(wb.get_cell_value("sheet2", "A2"),
                         wb.get_cell_value("sheet2", "A3"))

    def test_second_dollar_sheet_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet2", "A2", "=sheet1!A1")
        wb.set_cell_contents("sheet2", "A3", "=sheet1!A$1")
        self.assertEqual(wb.get_cell_value("sheet2", "A2"),
                         wb.get_cell_value("sheet2", "A3"))

    def test_double_dollar_sheet_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1")
        wb.set_cell_contents("sheet2", "A2", "=sheet1!A1")
        wb.set_cell_contents("sheet2", "A3", "=sheet1!$A$1")
        self.assertEqual(wb.get_cell_value("sheet2", "A2"),
                         wb.get_cell_value("sheet2", "A3"))

    def test_invalid_dollar_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=$A$2$")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.PARSE_ERROR)


if __name__ == "__main__":
    unittest.main()
