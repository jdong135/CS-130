"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
from context import sheets


class BooleanTests(unittest.TestCase):
    """
    Tests for booleans in lark module.
    """

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

    def test_concat_str_into_bool(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '\'tru')
        wb.set_cell_contents('sheet1', 'b1', '=A1 & "e"')
        self.assertEqual(wb.get_cell_value('sheet1', 'b1'), "true")

    def test_equality_of_diff_types(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=True == "True"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)
        wb.set_cell_contents('sheet1', 'a1', '=True = "True"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)

    def test_inequality_of_diff_types(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=True != "True"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)
        wb.set_cell_contents('sheet1', 'a1', '=True <> "True"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_blue_string_equality(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '="BLUE" = "blue"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)
        wb.set_cell_contents('sheet1', 'a1', '="BLUE" < "blue"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)
        wb.set_cell_contents('sheet1', 'a1', '="BLUE" > "blue"')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)

    def test_ascii_inequality(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '="a" < "["')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)

    def test_false_less_than_true(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=False < True')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_basic_number_inequality(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=3.0 == 3')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)
        wb.set_cell_contents('sheet1', 'a1', '=3.0 > 3.1')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)
        wb.set_cell_contents('sheet1', 'a1', '=3 < 5')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_string_greater_than_number(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '="12" > 12')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_string_greater_than_number(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '="TRUE" > False')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), False)
        wb.set_cell_contents('sheet1', 'b1', '=False < "TRUE"')
        self.assertEqual(wb.get_cell_value('sheet1', 'b1'), False)

    def test_booleans_bigger_than_nums(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'a1', '=False > 5')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)


if __name__ == "__main__":
    unittest.main()
