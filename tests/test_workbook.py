"""
Test for implementation of Workbook
"""

import unittest
import os
import sys
import string
import random

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import *

MAX_SHEETS_TEST = 100000
MAX_STR_LEN_TEST = 8
INVALID_CHARS = ["¿", "\"", "░", "Ä", "ô"]


class Workbook_New_Sheet(unittest.TestCase):
    def test_new_sheet1(self):
        wb = Workbook()
        wb.new_sheet()
        self.assertEqual(list(wb.spreadsheets.keys())[0], 'sheet1')

    def test_new_sheet2(self):
        wb = Workbook()
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
        wb = Workbook()
        for _ in range(MAX_SHEETS_TEST):
            wb.new_sheet()
        sheet_names = list(wb.spreadsheets.keys())
        for i in range(MAX_SHEETS_TEST):
            self.assertEqual(sheet_names[i], f"sheet{str(i + 1)}")

    def test_new_sheet_invalid_char(self):
        wb = Workbook()
        ch = INVALID_CHARS[random.randint(0, len(INVALID_CHARS) - 1)]
        sheet_name = list(''.join(random.choices(
            string.ascii_lowercase, k=MAX_STR_LEN_TEST)))
        sheet_name[random.randint(0, MAX_STR_LEN_TEST - 1)] = ch
        with self.assertRaises(ValueError):
            wb.new_sheet(''.join(sheet_name))

    def test_new_sheet_empty_string(self):
        wb = Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet("")

    def test_new_sheet_white_space(self):  # failing
        wb = Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet(" ")

    def test_new_sheet_head_white_space(self):
        wb = Workbook()
        sheet_name = list(''.join(random.choices(
            string.ascii_lowercase, k=MAX_STR_LEN_TEST)))
        sheet_name[0] = " "
        with self.assertRaises(ValueError):
            wb.new_sheet(''.join(sheet_name))


if __name__ == "__main__":
    unittest.main()
