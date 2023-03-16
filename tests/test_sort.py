"""
Unit tests for implementation of Workbook.sort_region()
"""
import unittest
import decimal
from context import sheets
from utils import block_equal, store_stdout, restore_stdout, on_cells_changed, sort_notify_list


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

    def test_sort_single_cell_with_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "3")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'B1', [1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal("3"))
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal("3"))
        self.assertEqual([], output)

    def test_sort_ref_to_empty(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "=3")
        wb.set_cell_contents("sheet1", "A2", "=B2")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'A2', [1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        self.assertEqual(wb.get_cell_contents("sheet1", "A1"), "=B1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal("0"))
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), decimal.Decimal("3"))
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'"]
        self.assertEqual(expected, output)

    def test_sort_with_functions(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "=IF(B1, 1, C1)")
        wb.set_cell_contents("sheet1", "B1", "TRUE")
        wb.set_cell_contents("sheet1", "C1", "2")
        wb.set_cell_contents("sheet1", "A2", "=IF(B2, 1, C2)")
        wb.set_cell_contents("sheet1", "B2", "FALSE")
        wb.set_cell_contents("sheet1", "C2", "=1/0")
        wb.set_cell_contents("sheet1", "A3", "=IF(B3, 1, C3)")
        wb.set_cell_contents("sheet1", "B3", "TRUE")
        wb.set_cell_contents("sheet1", "C3", "=3")
        wb.set_cell_contents("sheet1", "A4", "=IF(B4, 1, C4)")
        wb.set_cell_contents("sheet1", "B4", "FALSE")
        wb.set_cell_contents("sheet1", "C4", "'STRING")
        #     |  A  |  B  |  C  |
        #  1  |    1| TRUE|    2|
        #  2  |#DIV/|FALSE|#DIV/|
        #  3  |    1| TRUE|    3|
        #  4  |STRIN|FALSE|STRIN|
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'C4', [1, 2, -3])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        zero_div = sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO
        block_equal(wb, self,
                    [zero_div, False, zero_div,
                     1, True, 3,
                     1, True, 2,
                     "STRING", False, "STRING"])
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'",
                    "'Sheet1', 'B1'", "'Sheet1', 'B2'", 
                    "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet1', 'C3'"]
        self.assertEqual(expected, output)

    def test_sort_error_types(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        new_stdo, sys_out = store_stdout()
        wb.set_cell_contents("sheet1", "A1", "#DIV/0!")
        wb.set_cell_contents("sheet1", "A2", "#VALUE!")
        wb.set_cell_contents("sheet1", "A3", "#REF!")
        wb.set_cell_contents("sheet1", "A4", "#NAME?")
        wb.set_cell_contents("sheet1", "A5", "#CIRCREF!")
        wb.set_cell_contents("sheet1", "A6", "#ERROR!")
        wb.notify_cells_changed(on_cells_changed)
        wb.sort_region('Sheet1', 'A1', 'A6', [1])
        output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        self.assertEqual(wb.get_cell_value("sheet1", "A1"
        ).get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"
        ).get_type(), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"
        ).get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "A4"
        ).get_type(), sheets.cell_error.CellErrorType.BAD_NAME)
        self.assertEqual(wb.get_cell_value("sheet1", "A5"
        ).get_type(), sheets.cell_error.CellErrorType.TYPE_ERROR)
        self.assertEqual(wb.get_cell_value("sheet1", "A6"
        ).get_type(), sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'A5'", "'Sheet1', 'A6'"]
        self.assertEqual(expected, output)


if __name__ == "__main__":
    unittest.main()
