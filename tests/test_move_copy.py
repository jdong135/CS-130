"""
Unit tests for implementation of Workbook.move_cells() and Workbook.copy_cells()
"""
import unittest
import decimal
from context import sheets
from utils import store_stdout, restore_stdout, sort_notify_list, on_cells_changed


class WorkbookMoveCopyCells(unittest.TestCase):
    """
    Unit tests for Workbook.move_cells
    """

    # def test_basic_move1(self):
    #     """
    #     Basic Move to the right using top left and bottom right corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "A1", "C2", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move2(self):
    #     """
    #     Basic Move to the right using bottom right and top left corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "C2", "A1", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move3(self):
    #     """
    #     Basic Move right using top right and bottom left corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "C1", "A2", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move4(self):
    #     """
    #     Basic Move right using bottom left and top right corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "A2", "C1", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move5(self):
    #     """
    #     Basic Move down using top left and bottom right corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "A1", "C2", "A3")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move6(self):
    #     """
    #     Basic Move down using bottom right and top left corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "C2", "A1", "A3")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move7(self):
    #     """
    #     Basic Move down using bottom left and top right corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "A2", "C1", "A3")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move8(self):
    #     """
    #     Basic Move down using top right and bottom left corners.
    #     """
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sheet1", "C1", "A2", "A3")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_case_insensitive(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.move_cells("sHeET1", "A1", "C2", "A3")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_basic_move_to_diff_sheet(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "C1", "A2", "A3", "sheet2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "A3"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "B3"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "C3"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "A4"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "B4"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet2", "C4"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'B1'", "'Sheet1', 'B2'",
    #                 "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet2', 'A3'", "'Sheet2', 'A4'",
    #                 "'Sheet2', 'B3'", "'Sheet2', 'B4'", "'Sheet2', 'C3'", "'Sheet2', 'C4'"]
    #     self.assertEqual(expected, output)

    def test_move_block_with_empty_cells(self):
        # new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "1")
        wb.set_cell_contents("sheet1", "A2", "2")
        wb.set_cell_contents("sheet1", "B3", "3")
        wb.notify_cells_changed(on_cells_changed)
        wb.move_cells("sheet1", "A3", "B1", "B1", "sheet1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B1"), decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B2"), decimal.Decimal(2))
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "B3"), None)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "C1"), None)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "C2"), None)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "C3"), decimal.Decimal(3))
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1"), None)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2"), None)
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A3"), None)
        # output = sort_notify_list(restore_stdout(new_stdo, sys_out))
        expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'B1'", "'Sheet1', 'B2'",
                    "'Sheet1', 'B3'", "'Sheet1', 'C3'"]
        # self.assertEqual(expected, output)

    # def test_basic_sheet_name_not_found(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.notify_cells_changed(on_cells_changed)
    #     with self.assertRaises(KeyError):
    #         wb.move_cells("sheet3", "A1", "C2", "A3")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     self.assertEqual([], output)

    # def test_basic_cell_invalid(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     with self.assertRaises(ValueError):
    #         wb.move_cells("sheet1", "A1", "C2", "ZZZZZ99999")

    # def test_basic_overlap(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A2", "C1", "B2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B3"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C3"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D3"), "botright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'B1'", "'Sheet1', 'B2'",
    #                 "'Sheet1', 'B3'", "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet1', 'C3'",
    #                 "'Sheet1', 'D2'", "'Sheet1', 'D3'"]
    #     self.assertEqual(expected, output)

    # def test_basic_complete_overlap(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'topleft")
    #     wb.set_cell_contents("sheet1", "B1", "'topmid")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "'botright")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A2", "C1", "A1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), "topleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), "topmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), "botright")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     self.assertEqual([], output)

    # def test_formula_move1(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "= A1 * B1")
    #     wb.move_cells("sheet1", "A1", "C2", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "123")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), decimal.Decimal(2))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), decimal.Decimal(246))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_formula_move2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "= A1 * B1")
    #     wb.move_cells("sheet1", "C2", "A1", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "123")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), decimal.Decimal(2))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), decimal.Decimal(246))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_formula_move3(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "= A1 * B1")
    #     wb.move_cells("sheet1", "A2", "C1", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "123")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), decimal.Decimal(2))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), decimal.Decimal(246))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_formula_move4(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "'topright")
    #     wb.set_cell_contents("sheet1", "A2", "'botleft")
    #     wb.set_cell_contents("sheet1", "B2", "'botmid")
    #     wb.set_cell_contents("sheet1", "C2", "= A1 * B1")
    #     wb.move_cells("sheet1", "C1", "A2", "D1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D1"), "123")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E1"), decimal.Decimal(2))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F1"), "topright")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), "botleft")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), "botmid")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F2"), decimal.Decimal(246))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2"), None)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), None)

    # def test_formula_overwritten_by_string(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "1")
    #     wb.set_cell_contents("sheet1", "B1", "=A1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A1", "B1", "B1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1"), decimal.Decimal(1))
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1"), decimal.Decimal(1))
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "C1"), "=B1")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'B1'", "'Sheet1', 'C1'"]
    #     self.assertEqual(expected, output)

    # def test_absolute_row_addition(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "1")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 + B$1")
    #     wb.move_cells("sheet1", "A1", "C1", "B2")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "D2"), "=B2 + C$1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), decimal.Decimal(1))

    # def test_absolute_col_addition(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "1")
    #     wb.set_cell_contents("sheet1", "B1", "=3 + $A1")
    #     wb.move_cells("sheet1", "A1", "B1", "B2")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "C2"), "=3 + $A2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2"), decimal.Decimal(3))

    # def test_absolute_row_col_addition(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "10")
    #     wb.set_cell_contents("sheet1", "B1", "4")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 + $B$1")
    #     wb.move_cells("sheet1", "A1", "C1", "B2")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "D2"), "=B2 + $B$1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2"), decimal.Decimal(10))

    # def test_absolute_row_addition2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "1")
    #     wb.set_cell_contents("sheet1", "B1", "2")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 + B$1")
    #     wb.set_cell_contents("sheet1", "E1", "5")
    #     wb.move_cells("sheet1", "A1", "C1", "D3")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "F3"), "=D3 + E$1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F3"), decimal.Decimal(6))

    # def test_absolute_col_addition2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "1")
    #     wb.set_cell_contents("sheet1", "B1", "=3 + $A1")
    #     wb.set_cell_contents("sheet1", "A2", "5")
    #     wb.move_cells("sheet1", "A1", "B1", "D2")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "E2"), "=3 + $A2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "E2"), decimal.Decimal(8))

    # def test_absolute_row_col_addition2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "10")
    #     wb.set_cell_contents("sheet1", "B1", "4")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 + $B$1")
    #     wb.move_cells("sheet1", "A1", "C1", "D3")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "F3"), "=D3 + $B$1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "F3"), decimal.Decimal(10))

    # def test_div_zero_after_move(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "10")
    #     wb.set_cell_contents("sheet1", "B1", "5")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 / B$1")
    #     wb.move_cells("sheet1", "A1", "C1", "B2")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "D2"), "=B2 / C$1")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "D2").get_type(), sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    # def test_invalid_cell_ref(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 * B1")
    #     wb.set_cell_contents("sheet1", "A2", "4.5")
    #     wb.set_cell_contents("sheet1", "B2", "3.1")
    #     wb.set_cell_contents("sheet1", "C2", "=A2 * B2")
    #     wb.move_cells("sheet1", "C1", "C2", "B1")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "B1"), "=#REF! * A1")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "B2"), "=#REF! * A2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)

    # def test_parse_error_move(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=5Q sheet1!B1 Z")
    #     wb.set_cell_contents("sheet1", "A2", "=7N- B2 k")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A1").get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "A2").get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)
    #     wb.move_cells("sheet1", "A1", "A2", "C1")
    #     self.assertEqual(wb.get_cell_contents(
    #         "Sheet1", "C1"), "=5Q sheet1!B1 Z")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "C2"), "=7N- B2 k")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1").get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2").get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)

    # def test_move_sheet_ref(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=sheet2!A1")
    #     wb.move_cells("sheet1", "A1", "A1", "B1")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), "=sheet2!B1")

    # def test_mixed_ref(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=Sheet2!A1 + Sheet1!A2 + A3")
    #     wb.move_cells("sheet1", "A1", "A1", "B1")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"),
    #                      "=Sheet2!B1 + Sheet1!B2 + B3")

    # def test_ref_adjacent_char(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents(
    #         "sheet1", "A1", "=1 + (A2) + (Sheet1!A3) + (Sheet2!A4)")
    #     wb.move_cells("sheet1", "A1", "A1", "B1")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), "=1 + (B2) + (Sheet1!B3) "
    #                      "+ (Sheet2!B4)")

    # def test_basic_move_noitfy(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=5")
    #     wb.set_cell_contents("sheet1", "A2", "=5 + A1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A1", "A2", "B1")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'",
    #                 "'Sheet1', 'B1'", "'Sheet1', 'B2'"]
    #     self.assertEqual(expected, output)

    # def test_copy_no_overlap(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1*B1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.copy_cells("sheet1", "A1", "C1", "A2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A2"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C2"), "=A2*B2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C1"), "=A1*B1")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A2'", "'Sheet1', 'B2'", "'Sheet1', 'C2'"]
    #     self.assertEqual(expected, output)

    # def test_move_no_overlap(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1*B1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A1", "C1", "A2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A2"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C2"), "=A2*B2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C1"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'", "'Sheet1', 'B1'",
    #                 "'Sheet1', 'B2'", "'Sheet1', 'C1'", "'Sheet1', 'C2'"]
    #     self.assertEqual(expected, output)

    # def test_copy_no_overlap_shift(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1*B1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.copy_cells("sheet1", "A1", "C1", "B2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C2"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "D2"), "=B2*C2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C1"), "=A1*B1")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'B2'", "'Sheet1', 'C2'", "'Sheet1', 'D2'"]
    #     self.assertEqual(expected, output)

    # def test_move_no_overlap_shift(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "'123")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1*B1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A1", "C1", "B2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "'123")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C2"), "5.3")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "D2"), "=B2*C2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), None)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C1"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'B1'", "'Sheet1', 'B2'",
    #                 "'Sheet1', 'C1'", "'Sheet1', 'C2'", "'Sheet1', 'D2'"]
    #     self.assertEqual(expected, output)

    # def test_copy_invalid_cell_ref(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "C1", "=A1 * B1")
    #     wb.set_cell_contents("sheet1", "A2", "4.5")
    #     wb.set_cell_contents("sheet1", "B2", "3.1")
    #     wb.set_cell_contents("sheet1", "C2", "=A2 * B2")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.copy_cells("sheet1", "C1", "C2", "B1")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "B1"), "=#REF! * A1")
    #     self.assertEqual(wb.get_cell_contents("Sheet1", "B2"), "=#REF! * A2")
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "B2").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_value(
    #         "Sheet1", "C2").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'B1'", "'Sheet1', 'B2'"]
    #     self.assertEqual(expected, output)

    # def test_move_to_invalid(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "A2", "4.5")
    #     wb.set_cell_contents("sheet1", "B2", "3.1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     with self.assertRaises(ValueError):
    #         wb.move_cells("sheet1", "A1", "C2", "ZZZZ9999")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     self.assertEqual([], output)

    # def test_copy_to_invalid(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "5.3")
    #     wb.set_cell_contents("sheet1", "A2", "4.5")
    #     wb.set_cell_contents("sheet1", "B2", "3.1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     with self.assertRaises(ValueError):
    #         wb.copy_cells("sheet1", "A1", "C2", "ZZZZ9999")
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     self.assertEqual([], output)

    # def test_move_zero_error(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "0")
    #     wb.set_cell_contents("sheet1", "A2", "1")
    #     wb.set_cell_contents("sheet1", "B1", '= 1 / A1')
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "B1", "B1", "B2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2"),
    #                      decimal.Decimal(1))
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'B1'", "'Sheet1', 'B2'"]
    #     self.assertEqual(expected, output)

    # def test_move_ref_error(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "=#REF!*A1")
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "A1", "B1", "A2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "=#REF!*A2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2").get_type(),
    #                      sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), None)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), None)
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'A1'", "'Sheet1', 'A2'",
    #                 "'Sheet1', 'B1'", "'Sheet1', 'B2'"]
    #     self.assertEqual(expected, output)

    # def test_copy_ref_error(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "2.2")
    #     wb.set_cell_contents("sheet1", "B1", "=#REF!*A1")
    #     wb.copy_cells("sheet1", "A1", "B1", "A2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B2"), "=#REF!*A2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2").get_type(),
    #                      sheets.cell_error.CellErrorType.BAD_REFERENCE)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "A1"), "2.2")
    #     self.assertEqual(wb.get_cell_contents("sheet1", "B1"), "=#REF!*A1")

    # def test_move_fix_circ_error(self):
    #     new_stdo, sys_out = store_stdout()
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=A2")
    #     wb.set_cell_contents("sheet1", "A2", "=B2")
    #     wb.set_cell_contents("sheet1", "B2", "=B1")
    #     wb.set_cell_contents("sheet1", "B1", "=A1")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     wb.notify_cells_changed(on_cells_changed)
    #     wb.move_cells("sheet1", "B2", "B2", "C2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(0))
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2"), decimal.Decimal(0))
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(0))
    #     self.assertEqual(wb.get_cell_value("sheet1", "C2"), decimal.Decimal(0))
    #     output = sort_notify_list(restore_stdout(new_stdo, sys_out))
    #     expected = ["'Sheet1', 'B2'", "'Sheet1', 'C2'"]
    #     self.assertEqual(expected, output)

    # def test_copy_fix_circ_error(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=A2")
    #     wb.set_cell_contents("sheet1", "A2", "=B2")
    #     wb.set_cell_contents("sheet1", "B2", "=B1")
    #     wb.set_cell_contents("sheet1", "B1", "=A1")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     wb.copy_cells("sheet1", "B2", "B2", "C2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C2"),
    #                      "=C1")

    # def test_move_fix_circ_error2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=B1")
    #     wb.set_cell_contents("sheet1", "B1", "=C1")
    #     wb.set_cell_contents("sheet1", "C1", "=A1")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "C1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     wb.move_cells("sheet1", "C1", "C1", "C2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1"),
    #                      decimal.Decimal(0))
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1"),
    #                      decimal.Decimal(0))
    #     self.assertEqual(wb.get_cell_contents("sheet1", "C1"),
    #                      None)
    #     self.assertEqual(wb.get_cell_value("sheet1", "C2"),
    #                      decimal.Decimal(0))

    # def test_move_fix_circ_error3(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("sheet1", "A1", "=B1")
    #     wb.set_cell_contents("sheet1", "B1", "=C1")
    #     wb.set_cell_contents("sheet1", "C1", "=A1")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "C1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     wb.copy_cells("sheet1", "C1", "C1", "C2")
    #     self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "B1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "C1").get_type(),
    #                      sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertEqual(wb.get_cell_value("sheet1", "C2"),
    #                      decimal.Decimal(0))

    # def test_move_diff_sheet_ref(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents('Sheet2', 'A1', '5')
    #     wb.set_cell_contents('sheet1', 'A1', '=Sheet2!A1')
    #     wb.set_cell_contents('sheet1', 'A2', '=A1 * 5')
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2"),
    #                      decimal.Decimal(25))
    #     wb.set_cell_contents('Sheet2', 'B1', '10')
    #     wb.move_cells('sheet1', 'A1', 'A2', 'B1')
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2"),
    #                      decimal.Decimal(50))
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A1"), None)
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A2"), None)

    # def test_move_absolute_diff_sheet(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents('Sheet2', 'A1', '5')
    #     wb.set_cell_contents('sheet1', 'A1', '=Sheet2!$A$1')
    #     wb.set_cell_contents('sheet1', 'A2', '=A1 * 5')
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2"),
    #                      decimal.Decimal(25))
    #     wb.move_cells('sheet1', 'A1', 'A2', 'B1')
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2"),
    #                      decimal.Decimal(25))
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A1"), None)
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A2"), None)

    # def test_move_absolute_diff_sheet2(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents('Sheet2', 'A1', '5')
    #     wb.set_cell_contents('sheet1', 'A1', '=Sheet2!$A1')
    #     wb.set_cell_contents('sheet1', 'A2', '=A1 * 5')
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2"),
    #                      decimal.Decimal(25))
    #     wb.move_cells('sheet1', 'A1', 'A2', 'B1')
    #     self.assertEqual(wb.get_cell_value("sheet1", "B2"),
    #                      decimal.Decimal(25))
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A1"), None)
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A2"), None)

    # def test_move_absolute_diff_sheet3(self):
    #     wb = sheets.Workbook()
    #     wb.new_sheet()
    #     wb.new_sheet()
    #     wb.set_cell_contents('Sheet2', 'A1', '5')
    #     wb.set_cell_contents('sheet1', 'A1', '=Sheet2!A$1')
    #     wb.set_cell_contents('sheet1', 'A2', '=A1 * 5')
    #     self.assertEqual(wb.get_cell_value("sheet1", "A2"),
    #                      decimal.Decimal(25))
    #     wb.move_cells('sheet1', 'A1', 'A2', 'A3')
    #     self.assertEqual(wb.get_cell_value("sheet1", "A4"),
    #                      decimal.Decimal(25))
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A1"), None)
    #     self.assertEqual(wb.get_cell_contents('sheet1', "A2"), None)


if __name__ == "__main__":
    unittest.main()
