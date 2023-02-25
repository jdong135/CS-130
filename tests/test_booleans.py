"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
from context import sheets


class BooleanTests(unittest.TestCase):
    def test_basic_bool_inputs(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', 'true')
        assert isinstance(wb.get_cell_value('Sheet1', 'A1'), bool)
        assert wb.get_cell_value('Sheet1', 'A1')
        wb.set_cell_contents('Sheet1', 'A2', '=FaLsE')
        assert isinstance(wb.get_cell_value('Sheet1', 'A2'), bool)
        assert not wb.get_cell_value('Sheet1', 'A2')

    def test_equality_with_math(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=34 = 3.4 * 10')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

    def test_double_equality_with_math(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=34 == 3.4 * 10')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

    def test_chained_inequality(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'b1', '3')
        wb.set_cell_contents('sheet1', 'c1', '4')
        wb.set_cell_contents('sheet1', 'd1', '5')
        wb.set_cell_contents('sheet1', 'A1', '=B1 > C1 <= D1')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), False)

    def test_inequality_error_update(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=a1')
        wb.set_cell_contents('sheet1', 'b1', '=a1 < 5')
        wb.set_cell_contents('sheet1', 'a1', '=a2')
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), True)

    # Unsure what correct behavior is
    # def test_concat_str_into_bool(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents('sheet1', 'a1', '\'tru')
    #     wb.set_cell_contents('sheet1', 'b1', '=A1 & "e"')


if __name__ == "__main__":
    unittest.main()
