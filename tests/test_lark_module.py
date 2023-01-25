"""
Test for implementation of Lark Module formula Evaluator
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
from sheets import Workbook, lark_module, cell_error  # noqa
import logging

# Create and configure logger

logging.basicConfig(filename="logs/lark_module.log",

                    format='%(asctime)s %(message)s',

                    filemode='w')

# Creating an object

logger = logging.getLogger()

# Setting the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)


class Lark_Module_Basic(unittest.TestCase):
    def test_single_dec(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=1")
        self.assertEqual(value, 1)

    def test_single_str(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=\"str\"")
        self.assertEqual(value, "str")

    def test_basic1(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=1 + 7 - 124 * 7 / 14")
        self.assertEqual(value, -54)

    def test_basic2(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=2 * 6 - 4 + 2 / 2")
        self.assertEqual(value, 9)

    def test_basic_parens(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 - 7 + (4 * 8 - 7) - 2 / 2 + (7 - 4)")
        self.assertEqual(value, 24)

    def test_nested_parens(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 / 2 * (7 + 8 - (6 + (2 * 4)) - 4) + 7")
        self.assertEqual(value, 1)

    def test_negative1(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1")
        self.assertEqual(value, -1)

    def test_negative2(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1 + 2 * (-2 - 2 / (1 * -4))")
        self.assertEqual(value, -4)

#     def test_cell_ref1(self):
#         wb = Workbook()
#         wb.new_sheet("Sheet1")
#         wb.set_cell_contents("Sheet1", "A1", "=4")
#         wb.set_cell_contents("Sheet1", "A2", "=A1 + 1")
#         _, value = lark_module.evaluate_expr(
#             wb, None, "Sheet1", "=A2")
#         self.assertEqual(value, 5)

#     def test_cell_ref2(self):
#         wb = Workbook()
#         wb.new_sheet("Sheet1")
#         wb.new_sheet("Sheet2")
#         wb.set_cell_contents("Sheet1", "A1", "=1")
#         wb.set_cell_contents("Sheet2", "C5", "=3")
#         wb.set_cell_contents("Sheet2", "E1", "=Sheet2!C5 + C5 + Sheet1!A1")
#         _, value = lark_module.evaluate_expr(
#             wb, None, "Sheet2", "=E1")
#         self.assertEqual(value, 7)

#     def test_div_zero1(self):
#         wb = Workbook()
#         wb.new_sheet()
#         eval, _ = lark_module.evaluate_expr(
#             wb, None, "sheet1", "= 2 * 8 / 0 + 1")
#         self.assertEqual(eval.cell_error.get_detail(), "divide by zero")

#     def test_div_zero2(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "=1")
#         wb.set_cell_contents("sheet1", "A2", "=0")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 / A2")
#         self.assertEqual(eval.cell_error.get_detail(), "divide by zero")

#     def test_string_arithmetic1(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'1")
#         wb.set_cell_contents("sheet1", "A2", "'hello world")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 + A2")
#         self.assertEqual(eval.cell_error.get_detail(), "string arithmetic")

#     def test_string_arithmetic2(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'goodbye world")
#         wb.set_cell_contents("sheet1", "A2", "'hello world")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 + A2")
#         self.assertEqual(eval.cell_error.get_detail(), "string arithmetic")

#     def test_string_arithmetic3(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'10")
#         wb.set_cell_contents("sheet1", "A2", "'2")
#         _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 + A2")
#         self.assertEqual(value, 12)

#     def test_string_arithmetic4(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'1")
#         wb.set_cell_contents("sheet1", "A2", "'hello world")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 * A2")
#         self.assertEqual(eval.cell_error.get_detail(), "string arithmetic")

#     def test_string_arithmetic5(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'goodbye world")
#         wb.set_cell_contents("sheet1", "A2", "'hello world")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 * A2")
#         self.assertEqual(eval.cell_error.get_detail(), "string arithmetic")

#     def test_string_arithmetic6(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'10")
#         wb.set_cell_contents("sheet1", "A2", "'2")
#         _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 * A2")
#         self.assertEqual(value, 20)

#     def test_string_arithmetic7(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'123")
#         wb.set_cell_contents("sheet1", "A2", "5.3")
#         _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 * A2")
#         self.assertEqual(value, decimal.Decimal('651.9'))

#     def test_string_arithmetic7(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'  123")
#         wb.set_cell_contents("sheet1", "A2", "5.3")
#         _, value = lark_module.evaluate_expr(wb, None, "sheet1", "=A1 * A2")
#         self.assertEqual(value, decimal.Decimal('651.9'))

#     def test_bad_reference1(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'10")
#         eval, _ = lark_module.evaluate_expr(wb, None, "sheet1", "=Sheet2!A1")
#         self.assertEqual(eval.cell_error.get_detail(), "sheet name not found")

#     def test_bad_reference2(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'10")
#         eval, _ = lark_module.evaluate_expr(
#             wb, None, "sheet1", "=Sheet1!A10000")
#         self.assertEqual(eval.cell_error.get_detail(), "invalid location")

#     def test_bad_reference3(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "'10")
#         eval, _ = lark_module.evaluate_expr(
#             wb, None, "sheet1", "=Sheet1!ZZZZZ9999")
#         self.assertEqual(eval.cell_error.get_detail(), "invalid location")

#     def test_parse_error1(self):
#         wb = Workbook()
#         wb.new_sheet()
#         eval, _ = lark_module.evaluate_expr(
#             wb, None, "sheet1", "=e71s18d")
#         self.assertEqual(eval.cell_error.get_detail(), "parse error")

#     def test_initialized_empty_cell_ref(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", 'A1', "=B1 + C1")

#         wb.set_cell_contents("sheet1", "B1", "=C1")
#         eval, value = lark_module.evaluate_expr(
#             wb, wb.spreadsheets['sheet1'].cells['A1'], "sheet1", wb.get_cell_contents("sheet1", "B1"))
#         self.assertEqual(value, 0)

#     def test_error_propagation1(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "abc")
#         wb.set_cell_contents("sheet1", "B1", "=A1 + 1")  # Value!
#         wb.set_cell_contents("sheet1", "C1", "=B1 + 1")
#         eval, value = lark_module.evaluate_expr(wb, None, "sheet1", "=C1")
#         self.assertEqual(value.get_detail(), "invalid operation")

#     def test_error_propagation2(self):
#         wb = Workbook()
#         wb.new_sheet()
#         wb.set_cell_contents("sheet1", "A1", "0")
#         wb.set_cell_contents("sheet1", "B1", "=1 / A1")  # Div/0!
#         wb.set_cell_contents("sheet1", "C1", "=B1 + 1")
#         eval, value = lark_module.evaluate_expr(wb, None, "sheet1", "=C1")
#         self.assertEqual(value.get_detail(), "divide by zero")


class NewTest(unittest.TestCase):
    def test_add(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=1+1")
        self.assertEqual(value, decimal.Decimal(2))

    def test_mult(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=8*2")
        self.assertEqual(value, decimal.Decimal(16))

    def test_neg(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=(-5 * 2) * -2")
        self.assertEqual(value, decimal.Decimal(20))

    def test_concat(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", '="abc" & "def" & "ghi"')
        self.assertEqual(value, "abcdefghi")

    def test_parens(self):
        wb = Workbook()
        wb.new_sheet()
        _, value = lark_module.evaluate_expr(
            wb, None, "sheet1", '=(1 * 3) + -1 * (3 * -1)')
        self.assertEqual(value, decimal.Decimal(6))

    def test_circ_ref(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        value = wb.get_cell_value('sheet1', 'B1')
        assert isinstance(value, cell_error.CellError)
        assert value.get_type() == cell_error.CellErrorType.CIRCULAR_REFERENCE

    def test_circ_ref_update(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        wb.set_cell_contents('sheet1', 'B1', None)
        value = wb.get_cell_value('sheet1', 'B1')
        self.assertEqual(value, decimal.Decimal(0))


if __name__ == "__main__":
    unittest.main()
