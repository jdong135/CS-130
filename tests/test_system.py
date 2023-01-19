"""
Full system tests for Sheets module. Uses the "import sheets" syntax as opposed
to "from sheets import *" for better representation of user experience
"""

import unittest
import os
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
import sheets  # noqa


class System_Tests(unittest.TestCase):
    def test_system1(self):
        MAX_NUM = 10000000
        SIZE = 9999
        TEST_SAMPLE = 500

        self.assertEqual(sheets.version, "1.0")

        wb = sheets.Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet("my numbers\n")
        wb.new_sheet("my numbers")

        nums = []
        with open("test-data/numbers.txt", "r") as f:
            nums = f.readlines()
            for i in range(SIZE):
                nums[i] = int(nums[i])

        # Write all numbers into the sheet and check max bounds
        for i in range(SIZE):
            wb.set_cell_contents("my numbers", f"A{i + 1}", f"{nums[i]}")
        with self.assertRaises(ValueError):
            wb.set_cell_contents("my numbers", f"A{10000}", "10")

        # Test for zero division error.
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents(
                "my numbers", f"C{i + 1}", f"={nums[i]}/(15 / 15 * 10 - (5 + 2 + 3))")
            self.assertEqual(wb.get_cell_value(
                "my numbers", f"C{i + 1}").get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # Test for Parse errors.
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents(
                "my numbers", f"D{i + 1}", f"=A{i + 1}B{i + 1}")
            self.assertEqual(wb.get_cell_value(
                "my numbers", f"D{i + 1}").get_type(), sheets.CellErrorType.PARSE_ERROR)

        # Test for automatic filling.
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents("my numbers", f"E{i + 1}", f"=F{i + 1}")
            self.assertEqual(
                int(wb.get_cell_value("my numbers", f"E{i + 1}")), 0)

        # Test for circular referencing
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents("my numbers", f"F{i + 1}", f"=E{i + 1}")
            self.assertEqual(wb.get_cell_value(
                "my numbers", f"E{i + 1}").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
            self.assertEqual(wb.get_cell_value(
                "my numbers", f"F{i + 1}").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # Test for error precedence
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents(
                "my numbers", f"G{i + 1}", f"=C{i + 1} + D{i + 1} + E{i + 1} + G{i + 1}")
            self.assertEqual(wb.get_cell_value(
                "my numbers", f"G{i + 1}").get_type(), sheets.CellErrorType.PARSE_ERROR)

        # Test autoclearing of cells
        for i in range(TEST_SAMPLE):
            wb.set_cell_contents("my numbers", f"A{i + 1}", None)
            self.assertEqual(wb.get_cell_contents(
                "my numbers", f"A{i + 1}"), None)

        # Test that sheet can be deleted
        wb.del_sheet("my numbers")
        self.assertEqual(wb.num_sheets(), 0)


if __name__ == "__main__":
    unittest.main()
