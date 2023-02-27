"""
Test for implementation of Lark Module formula Evaluator
"""

import unittest
from context import sheets


class FunctionTests(unittest.TestCase):
    def test_bad_func_name(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AD(1,1)")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_NAME)

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
        wb.set_cell_contents(
            "sheet1", "A5", "=AND(1, 1 + 1, AND(1, AND(1, 5 * 4)))")
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
        wb.set_cell_contents("sheet1", "A1", "=AND(TruE, tRUE, TRUE)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)

    def test_and_error1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(A1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_move_AND(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(A1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), False)
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), True)

    def test_move_AND_abs_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(A$1)")
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), False)

    def test_move_AND_diff_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(sheet1!A1)")
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet2')
        self.assertEqual(wb.get_cell_value('sheet2', 'B2'), True)

    def test_move_AND_ref_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=true")
        wb.set_cell_contents("sheet1", "B2", "=AND(A1)")
        wb.copy_cells('sheet1', 'B2', 'B2', 'A2', 'sheet1')
        value = wb.get_cell_value('sheet1', 'A2')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_break_AND_circref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(C1)")
        wb.set_cell_contents("sheet1", "B1", "=AND(A1)")
        wb.set_cell_contents("sheet1", "C1", "=AND(B1)")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value = wb.get_cell_value('sheet1', 'B1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value = wb.get_cell_value('sheet1', 'C1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.move_cells('sheet1', 'A1', 'A1', 'A2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), False)

    def test_AND_on_str(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(\"a\")")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_AND_invalid_num_args(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND()")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    # Unsure what expected behavior is
    # def test_AND_str_concat(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'tru")
    #     wb.set_cell_contents("sheet1", "B1", "=AND(A1 & \"e\")")

    def test_xor(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=XOR(1, 0)")
        wb.set_cell_contents("sheet1", "A2", "=XOR (1, 0,  1)")
        wb.set_cell_contents("sheet1", "A3", "=XOR(1, 0, 1, AND(1 , 0))")
        wb.set_cell_contents("sheet1", "A4", "=XOR (1, 0,  1, AND(1, XOR(1, 0, 1 , 1, 0), 1))")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), True)

    def test_or(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=OR(1, 0, 0)")
        wb.set_cell_contents("sheet1", "A2", "=OR(0, 0, 0)")
        wb.set_cell_contents("sheet1", "A3", "=OR(AND(True, true, 1, 5, -3), 0, FALSE)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), True)

    def test_not(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=NOT(0)")
        wb.set_cell_contents("sheet1", "A2", "=NOT(1)")
        wb.set_cell_contents("sheet2", "A3", "=OR(XOR(1, 0, True), 0, NOT(True))")
        wb.set_cell_contents("sheet1", "A3", "=NOT(sheet2!A3)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1",  "A3"), True)


if __name__ == "__main__":
    unittest.main()
