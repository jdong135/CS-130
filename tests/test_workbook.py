"""
Unit tests for implementation of sheets.Workbook
"""

import unittest
import string
import random
import decimal
from context import sheets
from utils import store_stdout, restore_stdout

MAX_SHEETS_TEST = 100
MAX_STR_LEN_TEST = 100
MAX_ROW_COL_SIZE = 100
INVALID_CHARS = ["¿", "\"", "░", "😀", "\n", "\t"]


class WorkbookNewSheet(unittest.TestCase):
    """
    Unit tests for Workbook.new_sheet
    """

    def test_num_sheets(self):
        wb = sheets.Workbook()
        for _ in range(MAX_SHEETS_TEST):
            wb.new_sheet()
        self.assertEqual(wb.num_sheets(), MAX_SHEETS_TEST)

    def test_new_sheet1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        self.assertEqual(list(wb.spreadsheets.keys())[0], 'sheet1')

    def test_new_sheet2(self):
        wb = sheets.Workbook()
        wb.new_sheet("S1")      # S1
        wb.new_sheet("S2")      # S2
        wb.new_sheet("Sheet1")  # Sheet1
        wb.new_sheet()          # Sheet2
        wb.new_sheet("Sheet4")  # Sheet4
        wb.new_sheet()          # Sheet3
        wb.new_sheet()          # Sheet5
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[0], "s1")
        self.assertEqual(sheet_names[1], "s2")
        self.assertEqual(sheet_names[2], "sheet1")
        self.assertEqual(sheet_names[3], "sheet2")
        self.assertEqual(sheet_names[4], "sheet4")
        self.assertEqual(sheet_names[5], "sheet3")
        self.assertEqual(sheet_names[6], "sheet5")

    def test_new_sheet3(self):
        wb = sheets.Workbook()
        for _ in range(MAX_SHEETS_TEST):
            wb.new_sheet()
        sheet_names = list(wb.spreadsheets.keys())
        for i in range(MAX_SHEETS_TEST):
            self.assertEqual(sheet_names[i], f"sheet{str(i + 1)}")

    def test_new_sheet_invalid_char(self):
        wb = sheets.Workbook()
        ch = INVALID_CHARS[random.randint(0, len(INVALID_CHARS) - 1)]
        sheet_name = list(''.join(random.choices(
            string.ascii_lowercase, k=MAX_STR_LEN_TEST)))
        sheet_name[random.randint(0, MAX_STR_LEN_TEST - 1)] = ch
        with self.assertRaises(ValueError):
            wb.new_sheet(''.join(sheet_name))

    def test_new_sheet_empty_string(self):
        wb = sheets.Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet("")

    def test_new_sheet_white_space(self):
        wb = sheets.Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet(" ")

    def test_new_sheet_head_white_space(self):
        wb = sheets.Workbook()
        sheet_name = list(''.join(random.choices(
            string.ascii_lowercase, k=MAX_STR_LEN_TEST)))
        sheet_name[0] = " "
        with self.assertRaises(ValueError):
            wb.new_sheet(''.join(sheet_name))

    def test_new_sheet_tail_white_space(self):
        wb = sheets.Workbook()
        sheet_name = list(''.join(random.choices(
            string.ascii_lowercase, k=MAX_STR_LEN_TEST)))
        sheet_name[-1] = " "
        with self.assertRaises(ValueError):
            wb.new_sheet(''.join(sheet_name))

    def test_new_sheet_tail_white_space2(self):
        wb = sheets.Workbook()
        sheet_name = "a "
        with self.assertRaises(ValueError):
            wb.new_sheet(sheet_name)

    def test_new_sheet_duplicate1(self):
        wb = sheets.Workbook()
        wb.new_sheet("Sheet1")
        with self.assertRaises(ValueError):
            wb.new_sheet("sheet1")

    def test_new_sheet_duplicate2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.new_sheet("sheet1")


class WorkbookGetSheetExtent(unittest.TestCase):
    """
    Unit tests for Workbook.get_sheet_extent
    """

    def test_single_cell_extent(self):
        wb = sheets.Workbook()
        wb.new_sheet("S1")
        wb.set_cell_contents("S1", "A1", "test")
        extent = wb.get_sheet_extent("S1")
        self.assertEqual(extent, (1, 1))

    def test_max_extent(self):
        wb = sheets.Workbook()
        for i in range(MAX_SHEETS_TEST):
            wb.new_sheet()
            wb.set_cell_contents(f"sheet{str(i + 1)}", "ZZZZ9999", "test")
        for i in range(MAX_SHEETS_TEST):
            extent = wb.get_sheet_extent(f"sheet{str(i + 1)}")
            self.assertEqual(extent, (475254, 9999))

    def test_extent_key_error(self):
        wb = sheets.Workbook()
        wb.new_sheet("S1")
        wb.set_cell_contents("S1", "A1", "test")
        with self.assertRaises(KeyError):
            wb.get_sheet_extent("S2")


class WorkbookSetCellContents(unittest.TestCase):
    """
    Unit tests for Workbook.set_cell_contents
    """

    def test_add_cell(self):
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", "'test")
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), "'test")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "test")

    def test_decimal1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "3.00")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(3))
        self.assertEqual(wb.get_cell_contents('sheet1', 'A1'), "3.00")

    def test_decimal2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "2")
        wb.set_cell_contents('sheet1', 'A2', "2.5")
        wb.set_cell_contents('sheet1', 'A3', "=A1 * A2")
        # # We want to make sure it's not stored as '5.0'
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), decimal.Decimal(5))

    def test_decimal3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', "'2.340")
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'A1'), '2.340')

    def test_case_insensitive(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=#DIV/0!")
        wb.set_cell_contents("sheet1", "A2", "=A1 + 5")
        wb.set_cell_contents("sheet1", "A3", "=B1")
        wb.set_cell_contents("sheet1", "A4", "=b1")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'),
                         wb.get_cell_value('sheet1', 'A4'))


class WorkbookLoadAndSave(unittest.TestCase):
    """
    Unit tests for load_workbook and Workbook.save_workbook
    """

    def test_load_basic(self):
        with open("test-data/mock_workbook.json", "r", encoding="utf8") as fp:
            wb = sheets.Workbook.load_workbook(fp)
            self.assertEqual(wb.get_cell_value("sheet2", "A1"),
                             decimal.Decimal('651.9'))
            self.assertEqual(wb.get_cell_value(
                "sheet2", "A2"), '"Hello, world"')
            with open("test-data/mock_workbook2.json", "w", encoding="utf8") as fpw:
                wb.save_workbook(fpw)


class WorkbookCopySheet(unittest.TestCase):
    """
    Unit tests for Workbook.copy_Sheet
    """

    def test_copy_basic(self):
        wb = sheets.Workbook()
        wb.new_sheet("SHeet1")
        wb.set_cell_contents("sheet1", "A1", "=3")
        wb.set_cell_contents("sheet1", "A2", "=-2")
        wb.set_cell_contents("sheet1", "A3", "=A2 + A1")
        wb.new_sheet("sHeet1_1")
        wb.new_sheet("sheET1_2")
        wb.new_sheet("shEet1_3")
        wb.new_sheet("sheet1_4")
        wb.new_sheet("Sheet1_5")
        wb.set_cell_contents("sheet1_5", "A4", "Hello World")
        wb.set_cell_contents("sheet1", "A5", "=sheet1_5!A4")
        name, index = wb.copy_sheet("sheet1")
        self.assertEqual(name, "sheet1_6")
        self.assertEqual(index, 6)
        self.assertEqual(wb.get_cell_value("sheet1_6", "A3"),
                         decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value("sheet1_6", "A5"), "Hello World")

    def test_copy_key_error(self):
        wb = sheets.Workbook()
        for _ in range(5):
            wb.new_sheet()
        with self.assertRaises(KeyError):
            wb.copy_sheet("sheet6")

    def test_rename_sheet1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=1 + 1")
        wb.rename_sheet("sheet1", "renamed_sheet")

    def test_copy_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=A3")
        wb.copy_sheet("sheet1")
        self.assertEqual(wb.get_cell_contents("sheet1_1", "A1"), "=A3")


class WorkbookMoveSheet(unittest.TestCase):
    """
    Unit tests for Workbook.move_sheet
    """

    def test_move_to_0(self):
        wb = sheets.Workbook()
        wb.new_sheet("SHeet1")
        wb.new_sheet("sheet2")
        wb.move_sheet('sheEt2', 0)
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[0], 'sheet2')

    def test_swap_move(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.move_sheet("sheet2", 0)
        wb.move_sheet("sheet1", 0)
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[0], "sheet1")
        self.assertEqual(sheet_names[1], "sheet2")
        wb.move_sheet("sheet1", 1)
        wb.move_sheet("sheet2", 1)
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[0], "sheet1")
        self.assertEqual(sheet_names[1], "sheet2")

    def test_move_to_middle(self):
        wb = sheets.Workbook()
        wb.new_sheet("SHeet1")
        wb.new_sheet("sheet2")
        wb.new_sheet("sheet3")
        wb.new_sheet("sheet4")
        wb.new_sheet("sheet5")
        wb.move_sheet('sheet4', 2)
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[2], 'sheet4')

    def test_move_to_end(self):
        wb = sheets.Workbook()
        wb.new_sheet("SHeet1")
        wb.new_sheet("sheet2")
        wb.new_sheet("sheet3")
        wb.new_sheet("sheet4")
        wb.new_sheet("sheet5")
        wb.new_sheet("sheet6")
        wb.move_sheet('sheet4', 5)
        sheet_names = list(wb.spreadsheets.keys())
        self.assertEqual(sheet_names[5], 'sheet4')

    def test_invalid_index(self):
        wb = sheets.Workbook()
        wb.new_sheet("SHeet1")
        wb.new_sheet("sheet2")
        wb.new_sheet("sheet3")
        with self.assertRaises(IndexError):
            wb.move_sheet('sheet3', 3)


class WorkbookRenameSheet(unittest.TestCase):
    """
    Unit tests for Workbook.rename_sheet
    """

    def test_basic_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet("Sheet 1")
        wb.new_sheet("Sheet 2")
        wb.rename_sheet("Sheet 2", "Renamed2")
        wb.rename_sheet("Sheet 1", "Renamed1")
        self.assertEqual(list(wb.spreadsheets.keys()),
                         ["renamed1", "renamed2"])

    def test_rename_ref1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet3", "A1", "=3")
        wb.set_cell_contents("Sheet2", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents(
            "Sheet1", "A2", "=A1 + 2 * Sheet2!A1 - 3 * Sheet3!A1")
        wb.rename_sheet("Sheet2", "SheetB")
        wb.rename_sheet("Sheet3", "SheetC")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2"), decimal.Decimal(-4))

    def test_rename_ref2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet2", "A1", "=4")
        wb.set_cell_contents("sheet1", "A1", "=sheet3!A1")
        wb.rename_sheet("sheet2", "sheet3")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(4))

    def test_quoted_rename_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet3", "A1", "=3")
        wb.set_cell_contents("Sheet2", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents(
            "Sheet1", "A2", "=A1 + 2 * 'Sheet2'!A1 - 3 * 'Sheet3'!A1")
        wb.rename_sheet("Sheet2", "SheetB")
        wb.rename_sheet("Sheet3", "SheetC")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A2"), decimal.Decimal(-4))

    def test_invalid_sheet_rename1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(KeyError):
            wb.rename_sheet("S2", "S3")

    def test_invalid_sheet_rename2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("sheet1", " a ")

    def test_invalid_sheet_rename3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("sheet1", " ")

    def test_invalid_sheet_rename4(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.rename_sheet("sheet2", "sheet4")
        with self.assertRaises(ValueError):
            wb.rename_sheet("sheet3", "sheet4")

    def test_invalid_sheet_rename5(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("sheet1", "a ")

    def test_invalid_sheet_rename6(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("sheet1", "' '")

    def test_unparsable_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet2", "A1", "=5")
        wb.set_cell_contents("sheet1", "A1", "=5Q sheet2!A1 Z")
        wb.rename_sheet("sheet2", "sheet3")
        self.assertEqual(wb.get_cell_contents(
            "sheet1", "A1"), "=5Q sheet2!A1 Z")

    def test_error_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet2", "A1", "=0")
        wb.set_cell_contents("sheet1", "A1", "=5/Sheet2!A1")
        wb.rename_sheet("sheet2", "sheet3")
        contents = wb.get_cell_contents("sheet1", "A1")
        value = wb.get_cell_value("sheet1", "A1")
        self.assertEqual(contents, "=5/sheet3!A1")
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    def test_parentheses_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet2", "A1", "=1")
        wb.set_cell_contents("sheet1", "A1", "=5 * sheet2!A1")
        wb.set_cell_contents(
            "sheet2", "A2", "=1 + (sheet1!A1 - 6 * (A1 * 1)) * -1")
        wb.set_cell_contents("sheet1", "A2", "=(((((sheet2!A2)-1)+1)))")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), decimal.Decimal(2))

    def test_strip_newname_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "='Sheet1'!A5 + 'Sheet2'!A6")
        wb.rename_sheet("sheet2", "SheetBla")
        self.assertEqual(wb.get_cell_contents(
            "sheet1", "A1"), "=Sheet1!A5 + SheetBla!A6")

    def test_strip_some_quotes_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "='Sheet1'!A5 + 'Sheet2'!A6")
        wb.rename_sheet("sheet2", "Sheet Space")
        self.assertEqual(wb.get_cell_contents(
            "sheet1", "A1"), "=Sheet1!A5 + 'Sheet Space'!A6")

    def test_mathematical_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet('a')
        wb.new_sheet('3*a')
        wb.set_cell_contents("a", "A1", "=3*a!B1")
        wb.rename_sheet("a", "b")
        self.assertEqual(wb.get_cell_contents("b", "A1"), "=3*b!B1")

    def test_ref_self_updated_rename(self):
        wb = sheets.Workbook()
        wb.new_sheet('a')
        wb.set_cell_contents("a", "A1", "='b'!A2")
        wb.rename_sheet('a', 'b')
        self.assertEqual(wb.get_cell_contents("b", "A1"), "='b'!A2")


class WorkbookNotifyCellsChanged(unittest.TestCase):
    """
    Unit tests for Workbook.notify_cells_changed
    """

    def test_basic_callable1(self):
        def on_cells_changed(workbook, changed_cells):
            '''
            This function gets called when cells change in the workbook that the
            function was registered on.  The changed_cells argument is an iterable
            of tuples; each tuple is of the form (sheet_name, cell_location).
            '''
            _ = workbook
            print(f'Cell(s) changed: {changed_cells}')
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=B1 + 1")
        wb.set_cell_contents("Sheet1", "B1", "=C1 + 1")
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("Sheet1", "C1", "=4")
        output = restore_stdout(new_stdo, sys_out)
        expected = "Cell(s) changed: [('Sheet1', 'C1'), ('Sheet1', 'B1'), ('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_non_callable(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        with self.assertRaises(TypeError):
            wb.notify_cells_changed("on_cells_changed")

    def test_continual_notify(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("Sheet1", "A1", "=B1 + 1")
        wb.set_cell_contents("Sheet1", "B1", "=C1 + 1")
        wb.set_cell_contents("Sheet1", "C1", "=4")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1')]\n" \
            "[('Sheet1', 'B1'), ('Sheet1', 'A1')]\n" \
            "[('Sheet1', 'C1'), ('Sheet1', 'B1'), ('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_clear_cell_notify(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("Sheet1", "A1", "=5")
        wb.set_cell_contents("Sheet1", "A1", "")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1')]\n[('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_rename_notify1(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", "=5")
        wb.notify_cells_changed(on_cells_changed)
        wb.rename_sheet("sheet1", "sheet2")
        output = restore_stdout(new_stdo, sys_out)
        self.assertEqual("", output)

    def test_rename_notify2(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.new_sheet("sheet2")
        wb.set_cell_contents("sheet1", "A1", "=sheet3!A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
        wb.notify_cells_changed(on_cells_changed)
        wb.rename_sheet("sheet2", "sheet3")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_rename_notify3(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet("sheet11")
        wb.new_sheet("sheet12")
        wb.set_cell_contents("sheet11", "A1", "=sheet12!A1")
        wb.notify_cells_changed(on_cells_changed)
        wb.rename_sheet("sheet12", "sheet13")
        output = restore_stdout(new_stdo, sys_out)
        self.assertEqual("", output)

    def test_rename_notify4(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=5Q sheet2!A1 Z")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.cell_error.CellErrorType.PARSE_ERROR)
        wb.notify_cells_changed(on_cells_changed)
        wb.rename_sheet("sheet2", "sheet3")
        wb.rename_sheet("sheet1", "sheet2")
        output = restore_stdout(new_stdo, sys_out)
        self.assertEqual("", output)

    def test_multiple_notify1(self):
        def on_cells_changed1(workbook, cells_changed):
            _ = workbook
            print(cells_changed)

        def on_cells_changed2(workbook, cells_changed):
            _ = workbook
            print(f"Updated: {cells_changed}")
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed1)
        wb.notify_cells_changed(on_cells_changed1)
        wb.notify_cells_changed(on_cells_changed2)
        wb.set_cell_contents("Sheet1", "A1", "=5")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1')]\n[('Sheet1', 'A1')]\nUpdated: [('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_multiple_notify2(self):
        def on_cells_changed1(workbook, cells_changed):
            _ = workbook
            print(cells_changed)

        def on_cells_changed2(workbook, cells_changed):
            _ = workbook
            _ = cells_changed
            raise ValueError
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed1)
        wb.notify_cells_changed(on_cells_changed2)
        wb.set_cell_contents("sheet1", "A1", "=1")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_tree_notify(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=0")
        wb.set_cell_contents("Sheet1", "B1", "=A1 + 1")
        wb.set_cell_contents("Sheet1", "B2", "=A1 + 1")
        wb.set_cell_contents("Sheet1", "C1", "=B1 + 2")
        wb.set_cell_contents("Sheet1", "C2", "=B1 + 2")
        wb.set_cell_contents("Sheet1", "C3", "=B2 + 2")
        wb.set_cell_contents("Sheet1", "C4", "=B2 + 2")
        wb.notify_cells_changed(on_cells_changed)
        wb.set_cell_contents("Sheet1", "A1", "=5")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1'), ('Sheet1', 'B1'), ('Sheet1', 'C1'), " \
            "('Sheet1', 'C2'), ('Sheet1', 'B2'), ('Sheet1', 'C3'), " \
            "('Sheet1', 'C4')]\n"
        self.assertEqual(expected, output)

    def test_delete_notify1(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=sheet2!A1")
        wb.set_cell_contents("Sheet1", "A2", "=sheet2!A1 + 5")
        wb.set_cell_contents("Sheet1", "A3", "=sheet2!A2")
        wb.notify_cells_changed(on_cells_changed)
        wb.del_sheet("sheet2")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1'), ('Sheet1', 'A2'), ('Sheet1', 'A3')]\n"
        self.assertEqual(expected, output)

    def test_delete_notify2(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=sheet2!A1")
        wb.set_cell_contents("Sheet2", "A1", "=sheet1!A1")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.notify_cells_changed(on_cells_changed)
        wb.del_sheet("sheet2")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('Sheet1', 'A1')]\n"
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
        self.assertEqual(expected, output)

    def test_new_sheet_notify(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", "=Sheet2!A1")
        wb.notify_cells_changed(on_cells_changed)
        wb.new_sheet("sheet2")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('sheet1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_copy_notify1(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", "=5")
        wb.notify_cells_changed(on_cells_changed)
        wb.copy_sheet("sheet1")
        output = restore_stdout(new_stdo, sys_out)
        expected = "[('sheet1_1', 'A1')]\n"
        self.assertEqual(expected, output)

    def test_copy_notify2(self):
        def on_cells_changed(workbook, cells_changed):
            _ = workbook
            print(cells_changed)
        new_stdo, sys_out = store_stdout()
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=sheet1_1!A1")
        self.assertEqual(wb.get_cell_value(
            "Sheet1", "A1").get_type(), sheets.cell_error.CellErrorType.BAD_REFERENCE)
        wb.notify_cells_changed(on_cells_changed)
        wb.copy_sheet("sheet1")
        output = restore_stdout(new_stdo, sys_out)
        self.assertEqual(wb.get_cell_value("Sheet1_1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        expected = "[('Sheet1', 'A1')]\n[('sheet1_1', 'A1'), ('Sheet1', 'A1')]\n"
        self.assertEqual(expected, output)


if __name__ == "__main__":
    unittest.main()
