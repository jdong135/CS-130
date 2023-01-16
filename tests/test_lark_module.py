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
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("=1")
        value = eval.visit(tree)
        self.assertEqual(value, 1)

    def test_single_str(self):
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("=\"str\"")
        value = eval.visit(tree)
        self.assertEqual(value, "str")

    def test_basic1(self):
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("=1 + 7 - 124 * 7 / 14")
        value = eval.visit(tree)
        self.assertEqual(value, -54)

    def test_basic2(self):
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("=2 * 6 - 4 + 2 / 2")
        value = eval.visit(tree)
        self.assertEqual(value, 9)

    def test_basic_parens(self):
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("= 4 - 7 + (4 * 8 - 7) - 2 / 2 + (7 - 4)")
        value = eval.visit(tree)
        self.assertEqual(value, 24)

    def test_nested_parens(self):
        eval = FormulaEvaluator()
        parser = lark.Lark.open('../sheets/formulas.lark', start='formula')
        tree = parser.parse("= 4 / 2 * (7 + 8 - (6 + (2 * 4)) - 4) + 7")
        value = eval.visit(tree)
        self.assertEqual(value, 1)


if __name__ == "__main__":
    unittest.main()
