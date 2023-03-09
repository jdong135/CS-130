"""
Unit tests for implementation of Workbook.sort_region()
"""
import unittest
import decimal
from context import sheets
from utils import store_stdout, restore_stdout, sort_notify_list, on_cells_changed


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
        wb.sort_region('Sheet1', 'A1', 'C4', [-1, -2])


if __name__ == "__main__":
    unittest.main()
