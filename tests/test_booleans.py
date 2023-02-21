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


if __name__ == "__main__":
    unittest.main()