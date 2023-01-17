"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
import os
import sys
import lark

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import *  # noqa


class Lark_Module_Basic(unittest.TestCase):
    def test_single_dec(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(wb, None, "sheet1", "=1")
        self.assertEqual(value, 1)

    def test_single_str(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(wb, None, "sheet1", "=\"str\"")
        self.assertEqual(value, "str")

    def test_basic1(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=1 + 7 - 124 * 7 / 14")
        self.assertEqual(value, -54)

    def test_basic2(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "=2 * 6 - 4 + 2 / 2")
        self.assertEqual(value, 9)

    def test_basic_parens(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 - 7 + (4 * 8 - 7) - 2 / 2 + (7 - 4)")
        self.assertEqual(value, 24)

    def test_nested_parens(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= 4 / 2 * (7 + 8 - (6 + (2 * 4)) - 4) + 7")
        self.assertEqual(value, 1)

    def test_negative1(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1")
        self.assertEqual(value, -1)

    def test_negative2(self):
        wb = Workbook()
        wb.new_sheet()
        value = lark_module.evaluate_expr(
            wb, None, "sheet1", "= -1 + 2 * (-2 - 2 / (1 * -4))")
        self.assertEqual(value, -4)

    def test_cell_ref1(self):
        wb = Workbook()
        wb.new_sheet("Sheet1")
        wb.set_cell_contents("Sheet1", "A1", "=4")
        wb.set_cell_contents("Sheet1", "A2", "=A1 + 1")
        value = lark_module.evaluate_expr(
            wb, None, "Sheet1", "=A2")
        self.assertEqual(value, 5)

    def test_cell_ref2(self):
        wb = Workbook()
        wb.new_sheet("Sheet1")
        wb.new_sheet("Sheet2")
        wb.set_cell_contents("Sheet1", "A1", "=1")
        wb.set_cell_contents("Sheet2", "C5", "=3")
        wb.set_cell_contents("Sheet2", "E1", "=Sheet2!C5 + C5 + Sheet1!A1")
        value = lark_module.evaluate_expr(
            wb, None, "Sheet2", "=E1")
        self.assertEqual(value, 7)

    # def test_div_zero1(self):
    #     wb = Workbook()
    #     wb.new_sheet()
    #     eval = FormulaEvaluator(wb, "sheet1")
    #     parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    #     tree = parser.parse("= 2 * 8 / 0 + 1")
    #     eval.visit(tree)


if __name__ == "__main__":
    unittest.main()
