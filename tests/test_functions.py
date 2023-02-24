"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
from context import sheets


class FunctionTests(unittest.TestCase):
    def test_and1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(1, 1, 1)")
        wb.set_cell_contents("sheet1", "A2", "=AND(1, 0, 1)")
        wb.set_cell_contents("sheet1", "A3", "=AND(0, 1, 0)")
        wb.set_cell_contents("sheet1", "A4", "=AND(0, 0, 0)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)        
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)        
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), False)

    def test_and2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(1, 1 + 1, 1)")
        wb.set_cell_contents("sheet1", "A2", "=AND(1, 1 - 1, 1)")
        wb.set_cell_contents("sheet1", "A3", "=AND(0, 1, A2)")
        wb.set_cell_contents("sheet1", "A4", "=AND(A3, 0, A2)")
        wb.set_cell_contents("sheet1", "A5", "=AND(1, 1 + 1, AND(1, AND(1, 5 * 4)))")
        wb.set_cell_contents("sheet1", "A6", "=AND(AND(1))")
        wb.set_cell_contents("sheet1", "A7", "=AND(AND(0))")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)        
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)        
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A5"), True)        
        self.assertEqual(wb.get_cell_value("sheet1", "A6"), True)        
        self.assertEqual(wb.get_cell_value("sheet1", "A7"), False)

    def test_and3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(A1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(), 
                         sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)   

if __name__ == "__main__":
    unittest.main()
