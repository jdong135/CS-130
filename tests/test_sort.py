"""
Unit tests for implementation of Workbook.sort_region()
"""
import unittest
import decimal
from context import sheets
from utils import block_equal, store_stdout, restore_stdout, on_cells_changed, sort_notify_list

import logging
logging.basicConfig(filename="logs/results.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

class WorkbookSortCells(unittest.TestCase):
    """
    Unit tests for Workbook.move_cells
    """

    def test_sort_integers(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
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
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [1, -2])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self, [7, 20, 6,
                               7, 8, 3,
                               7, 4, 8,
                               2000, 1, 3])
        expected = ["'Sheet1', 'B1'", "'Sheet1', 'B2'", "'Sheet1', 'C1'", "'Sheet1', 'C2'"]
        self.assertEqual(expected, output)

    def test_sort_str_and_ints_sort(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
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
        #     |   A  |  B  |  C  |
        #  1  |     7|    8|    3|
        #  2  |     7|   20|    6|
        #  3  |STRING|    4|    8|
        #  4  |  2000|    1|    3|
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [-1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self, ["STRING", 4, 8,
                               2000, 1, 3,
                               7, 8, 3,
                               7, 20, 6
                               ])
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'A3'", "'Sheet1', 'A4'", 
                    "'Sheet1', 'B1'", "'Sheet1', 'B2'", "'Sheet1', 'B3'", "'Sheet1', 'B4'", 
                    "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet1', 'C3'", "'Sheet1', 'C4'"]
        self.assertEqual(expected, output)

    def test_sort_uninitialized(self):
        # wb = sheets.Workbook()
        # wb.new_sheet()
        # wb.set_cell_contents("sheet1", "A2", "=A1")
        # logger.info(wb.get_cell_contents("sheet1", "A1"))
        # logger.info(wb.get_cell_value("sheet1", "A1"))
        # logger.info(wb.get_cell_contents("sheet1", "A2"))
        # logger.info(wb.get_cell_value("sheet1", "A2"))

        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "B2", "20")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "STRING")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "3")
        wb.set_cell_contents("sheet1", "A4", "2000")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        #     |  A  |  B  |  C  |
        #  1  |    7|    8|    3|
        #  2  |     |   20|    6|
        #  3  |STRIN|    4|    3|
        #  4  | 2000|    1|    3|
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self, [None, 20, 6,
                               7, 8, 3,
                               2000, 1, 3,
                               "STRING", 4, 3])
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'A3'", "'Sheet1', 'A4'", 
                    "'Sheet1', 'B1'", "'Sheet1', 'B2'", "'Sheet1', 'B3'", "'Sheet1', 'B4'", 
                    "'Sheet1', 'C1'", "'Sheet1', 'C2'"]
        self.assertEqual(expected, output)

    def test_sort_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
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
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [-1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self, [2000, 1, 3,
                               7, 8, 3,
                               7, 20, 6,
                               sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO, 4, 8])
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A3'", "'Sheet1', 'A4'", 
                    "'Sheet1', 'B1'", "'Sheet1', 'B2'", "'Sheet1', 'B3'", 
                    "'Sheet1', 'B4'", "'Sheet1', 'C2'", "'Sheet1', 'C3'", "'Sheet1', 'C4'"]
        self.assertEqual(expected, output)

    def test_sort_1_cell_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.set_cell_contents("sheet1", "B1", "7")
        wb.set_cell_contents("sheet1", "C1", "1")
        wb.set_cell_contents("sheet1", "A2", "6")
        wb.set_cell_contents("sheet1", "B2", "=B1")
        wb.set_cell_contents("sheet1", "C2", "1")
        wb.set_cell_contents("sheet1", "A3", "2")
        wb.set_cell_contents("sheet1", "B3", "9")
        wb.set_cell_contents("sheet1", "C3", "10")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C3', [-1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self,
                    [6, sheets.cell_error.CellErrorType.BAD_REFERENCE, 1,
                     3, 7, 1,
                     2, 9, 10])
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'B1'"]
        self.assertEqual(expected, output)

    def test_sort_on_col_with_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
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
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C3', [2])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self,
                    [3, 7, 1,
                     6, 7, 1,
                     2, 9, 10])
        self.assertEqual([], output)

    def test_sort_two_empty_cells(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "7")
        wb.set_cell_contents("sheet1", "B1", "8")
        wb.set_cell_contents("sheet1", "C1", "3")
        wb.set_cell_contents("sheet1", "B2", "=A2")
        wb.set_cell_contents("sheet1", "C2", "6")
        wb.set_cell_contents("sheet1", "A3", "STRING")
        wb.set_cell_contents("sheet1", "B3", "4")
        wb.set_cell_contents("sheet1", "C3", "8")
        wb.set_cell_contents("sheet1", "B4", "1")
        wb.set_cell_contents("sheet1", "C4", "3")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [1])
        block_equal(wb, self,
                    [None, 0, 6,
                     None, 1, 3,
                     7, 8, 3,
                    "STRING", 4, 8])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A3'", "'Sheet1', 'A4'", 
                    "'Sheet1', 'B1'", "'Sheet1', 'B2'", "'Sheet1', 'B3'", "'Sheet1', 'B4'", 
                    "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet1', 'C3'", "'Sheet1', 'C4'"]
        self.assertEqual(expected, output)

    def test_sort_single_cell(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'A1', [1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        block_equal(wb, self, [3])
        self.assertEqual([], output)


if __name__ == "__main__":
    unittest.main()
