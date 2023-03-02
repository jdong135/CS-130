"""
Performance Analysis For Project 4
"""

import unittest
import cProfile
import pstats
import re  # pylint: disable=unused-import
from context import sheets


def make_test_json_files(cols, rows):
    wb = sheets.Workbook()
    wb.new_sheet("sheet1")
    for y in range(1, cols):
        c = sheets.string_conversions.num_to_col(y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"This is {c}{i}")
    with open(f"test-data/stresstest_load_{cols}_{rows}.json", "w",
            encoding="utf8") as fpw:
        wb.save_workbook(fpw)


def stress_load_workbook(self, f, outputname):
    pc = cProfile.Profile()
    pc.enable()
    with open(f, "r", encoding="utf8") as fp:
        wb = sheets.Workbook.load_workbook(fp)
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A1"), "This is A1")
    self.assertEqual(wb.get_cell_value("sheet1", "A3"), "This is A3")
    self.assertEqual(wb.get_cell_value("sheet1", "B1"), "This is B1")
    pc.dump_stats(f'logs/test_load_workbook_{outputname}.stats')
    with open(f'logs/test_load_workbook_stats_{outputname}.stats', 'w', encoding="utf8") as stream:
        p = pstats.Stats(
            f'logs/test_load_workbook_{outputname}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def stress_move(self, rows, cols):
    pc = cProfile.Profile()
    wb = sheets.Workbook()
    wb.new_sheet("sheet1")
    for y in range(1, cols):
        c = sheets.string_conversions.num_to_col(y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"This is {c}{i}")
    pc.enable()
    wb.move_cells("sheet1", "A1", f"{c}{i}", "B2")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "B2"), "This is A1")
    pc.dump_stats(f'logs/test_move_{rows}_{cols}.stats')
    with open(f'logs/test_move_stats_{rows}_{cols}.stats', 'w', encoding="utf8") as stream:
        p = pstats.Stats(
            f'logs/test_move_{rows}_{cols}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def stress_copy_sheet(self, rows, cols):
    pc = cProfile.Profile()
    wb = sheets.Workbook()
    wb.new_sheet("sheet1")
    for y in range(1, cols):
        c = sheets.string_conversions.num_to_col(y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"This is {c}{i}")
    pc.enable()
    _, name = wb.copy_sheet("sheet1")
    pc.disable()
    self.assertEqual(name, "sheet1_1")
    self.assertEqual(wb.get_cell_value("sheet1_1", "B2"), "This is B2")
    pc.dump_stats(f'logs/test_copy_sheet_{rows}_{cols}.stats')
    with open(f'logs/test_copy_sheet_stats_{rows}_{cols}.stats', 'w', encoding="utf8") as stream:
        p = pstats.Stats(
            f'logs/test_copy_sheet_{rows}_{cols}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()


class Bulk_Change_Contents_Tests(unittest.TestCase):
    """
    Initialize and execute all test cases. 
    """
    def test_load_workbook(self):
        stress_load_workbook(self, 'test-data/stresstest_load_50_2500.json', "1")
        stress_load_workbook(self, 'test-data/stresstest_load_100_2500.json', "2")
        stress_load_workbook(self, 'test-data/stresstest_load_200_2500.json', "3")

    def test_move(self):
        stress_move(self, 500, 10)
        stress_move(self, 1000, 10)
        stress_move(self, 1500, 10)

    def test_copy_sheet(self):
        stress_copy_sheet(self, 500, 10)
        stress_copy_sheet(self, 1000, 10)
        stress_copy_sheet(self, 1500, 10)


if __name__ == "__main__":
    unittest.main()
