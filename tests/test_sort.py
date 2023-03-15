"""
Unit tests for implementation of Workbook.sort_region()
"""
import unittest
import decimal
from context import sheets
from utils import block_equal


class WorkbookSortCells(unittest.TestCase):
    """
    Unit tests for Workbook.move_cells
    """

    def test_sort_integers(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "A2", "7")
        wb.set_cell_contents("sheet1", "B2", "20")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "7")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "8")
        wb.set_cell_contents("sheet1", "A4", "2000")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        #     |  A  |  B  |  C  |
        #  1  |    7|    8|    3|
        #  2  |    7|   20|    6|
        #  3  |    7|    4|    8|
        #  4  | 2000|    1|    3|
        wb.sort_region('Sheet1', 'A1', 'C4', [1, -2])
        block_equal(wb, self, [7,    20, 6,
                               7,    8,  3,
                               7,    4,  8,
                               2000, 1,  3])

    def test_sort_str_and_ints_sort(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "A2", "7")
        wb.set_cell_contents("sheet1", "B2", "20")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "STRING")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "8")
        wb.set_cell_contents("sheet1", "A4", "2000")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        #     |  A  |  B  |  C  |
        #  1  |    7|    8|    3|
        #  2  |    7|   20|    6|
        #  3  |    7|    4|    8|
        #  4  | 2000|    1|    3|
        wb.sort_region('Sheet1', 'A1', 'C4', [-1])

    def test_sort_uninitialized(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "B2", "20")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "STRING")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "8")
        wb.set_cell_contents("sheet1", "A4", "2000")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        #     |  A  |  B  |  C  |
        #  1  |    7|    8|    3|
        #  2  |     |    6|    6|
        #  3  |STRIN|    4|    8|
        #  4  | 2000|    1|    3|
        wb.sort_region('Sheet1', 'A1', 'C4', [1])

    def test_sort_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "A2", "7")
        wb.set_cell_contents("sheet1", "B2", "20")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "=1/0")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "8")
        wb.set_cell_contents("sheet1", "A4", "2000")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        wb.sort_region('Sheet1', 'A1', 'C4', [-1])
        block_equal(wb, self, [2000, 1, 3,
                               7,    8, 3,
                               7,    20, 6,
                               sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO, 4, 8])
    
    def test_sort_1_cell_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.set_cell_contents("sheet1", "B1", "7")
        wb.set_cell_contents("sheet1", "C1", "1")
        wb.set_cell_contents("sheet1", "A2", "6")
        wb.set_cell_contents("sheet1", "B2", "=B1")
        wb.set_cell_contents("sheet1", "C2", "1")
        wb.set_cell_contents("sheet1", "A3", "2")
        wb.set_cell_contents("sheet1", "B3", "9")
        wb.set_cell_contents("sheet1", "C3", "10")
        wb.sort_region('Sheet1', 'A1', 'C3', [-1])
        block_equal(wb, self, 
                            [6, sheets.cell_error.CellErrorType.BAD_REFERENCE, 1, 
                             3, 7, 1, 
                             2, 9, 10])


    def test_sort_on_col_with_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.set_cell_contents("sheet1", "B1", "7")
        wb.set_cell_contents("sheet1", "C1", "1")
        wb.set_cell_contents("sheet1", "A2", "6")
        wb.set_cell_contents("sheet1", "B2", "=B1")
        wb.set_cell_contents("sheet1", "C2", "1")
        wb.set_cell_contents("sheet1", "A3", "2")
        wb.set_cell_contents("sheet1", "B3", "9")
        wb.set_cell_contents("sheet1", "C3", "10")
        #     |  A  |  B  |  C  |
        #  1  |  3  |  7  |  1  |
        #  2  |  6  |  7  |  1  |
        #  3  |  2  |  9  | 10  |
        wb.sort_region('Sheet1', 'A1', 'C3', [2])
        block_equal(wb, self, 
                            [3, 7, 1, 
                            6, 7, 1, 
                            2, 9, 10])

    def test_sort_single_cell(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.sort_region('Sheet1', 'A1', 'A1', [1])
        block_equal(wb, self, [3])


if __name__ == "__main__":
    unittest.main()
