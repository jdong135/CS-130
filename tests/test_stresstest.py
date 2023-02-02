"""
Performance Analysis For Project 2
"""

import unittest
import os
import sys
import decimal
import cProfile
import re
import pstats

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import Workbook, lark_module, CellErrorType, CellError  # noqa

class CProfile_Test(unittest.TestCase):
    def test_cprofile_functionality(self):
        cProfile.run('re.compile("foo|bar")', 'logs/test_cprofile_functionality.stats')
        with open('logs/test_cprofile_functionality_stats.stats', 'w') as stream:
            p = pstats.Stats('logs/test_cprofile_functionality.stats', stream=stream)
            p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()
        cProfile.run('re.compile("foo|bar")')

def many_reference_one(self, rows):
    """
    Generates rows number of cells that reference cell A1. Then, updates
    A1's value. Records time to update A1's value
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.set_cell_contents("sheet1", "A1", "= 5000 * 2")
    for i in range(2, rows):
        wb.set_cell_contents("sheet1", f"A{i}", "=A1")
    pc.enable()
    wb.set_cell_contents("sheet1", "A1", "= 5000")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A2"), 5000)
    pc.dump_stats(f'logs/test_many_reference_one_{rows}.stats')
    with open(f'logs/test_many_reference_one_stats_{rows}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_many_reference_one_{rows}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def many_reference_many(self, rows: int , cols: int):
    """
    Generates two sheets. On sheet1, cells in a block sized row x cols
    reference the corresponding cell in sheet2. Then each cell in sheet2's
    value is updated.
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.new_sheet("sheet2")
    for y in range(cols):
        c = chr(65 + y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"= sheet2!{c}{i}")
    pc.enable()
    for y in range(cols):
        c = chr(65 + y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet2", f"{c}{i}", f"3")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A1"), 3)
    pc.dump_stats(f'logs/test_many_reference_many_{rows}_{cols}.stats')
    with open(f'logs/test_many_reference_many_stats_{rows}_{cols}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_many_reference_many_{rows}_{cols}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()


def chain(self, chain_size):
    """
    Generates a chain of cells that references the cells to the left. Updates
    the cell at the left end of the cell.
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.set_cell_contents("sheet1", "A1", "1")
    for i in range(2, chain_size):
        wb.set_cell_contents("sheet1", f"A{i}", f"= A{(i-1)}")
    pc.enable()
    wb.set_cell_contents("sheet1", "A1", "20")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", f"A{i}"), 
                     wb.get_cell_value("sheet1", "A1"))
    pc.dump_stats(f'logs/test_chain_{chain_size}.stats')
    with open(f'logs/test_chain_stats_{chain_size}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_chain_{chain_size}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def large_cycle(self, cycle_size):
    """
    Generates cycle_size number of cells that reference the cell to the right,
    creating a cycle of size cycle_size. Then, the value of all cells in the
    cycle are updated to a CellError of type CIRCULAR_REFERENCE
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    for i in range(2, cycle_size):
        wb.set_cell_contents("sheet1", f"A{i}", f"= A{(i-1)}")
    pc.enable()
    wb.set_cell_contents("sheet1", "A1", f"= A{cycle_size-1}")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(), CellErrorType.CIRCULAR_REFERENCE)
    pc.dump_stats(f'logs/test_large_cycle_{cycle_size}.stats')
    with open(f'logs/test_large_cycle_stats_{cycle_size}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_large_cycle_{cycle_size}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def delete_sheet1(self, rows, cols):
    """
    Generates two sheets. On sheet1, cells in a block sized row x cols
    reference the corresponding cell in sheet2. Then sheet2 is
    deleted.
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.new_sheet("sheet2")
    for y in range(cols):
        c = chr(65+y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"= sheet2!{c}{i}")
    pc.enable()
    wb.del_sheet("sheet2")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(), CellErrorType.BAD_REFERENCE)
    pc.dump_stats(f'logs/test_delete_sheet1_{rows}_{cols}.stats')
    with open(f'logs/test_delete_sheet1_stats_{rows}_{cols}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_delete_sheet1_{rows}_{cols}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def make_break_cycle(self, cycle_size, make_break):
    """
    Generates cycle_size number of cells that reference the cell to the right,
    creating a cycle of size cycle_size. Then, the value of all cells in the
    cycle are updated to a CellError of type CIRCULAR_REFERENCE. Repeats
    this process make_break number of times.
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    for i in range(2, cycle_size):
        wb.set_cell_contents("sheet1", f"A{i}", f"= A{(i-1)}")
    # make cycle
    pc.enable()
    # break cycle
    wb.set_cell_contents("sheet1", "A1", f"= A{cycle_size-1}")
    for _ in range(make_break):
        wb.set_cell_contents("sheet1", "A1", f"0")
        wb.set_cell_contents("sheet1", f"A1", f"= A{(cycle_size-1)}")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(), CellErrorType.CIRCULAR_REFERENCE)
    pc.dump_stats(f'logs/test_make_break_cycle_{cycle_size}_{make_break}.stats')
    with open(f'logs/test__make_break_cycle_stats_{cycle_size}_{make_break}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_make_break_cycle_{cycle_size}_{make_break}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def reference_rename_sheet(self, rows, cols):
    """
    Generates two sheets. On sheet1, cells in a block sized row x cols
    reference the corresponding cell in sheet2. Then sheet2 is
    renamed.
    """
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.new_sheet("sheet2")
    for y in range(cols):
        c = chr(65 + y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"= sheet2!{c}{i}")
    pc.enable()
    wb.rename_sheet("sheet2", "sheet3")
    pc.disable()
    for y in range(cols):
        c = chr(65 + y)
        for i in range(1, rows):
            wb.set_cell_contents("sheet3", f"{c}{i}", f"3")
    self.assertEqual(wb.get_cell_value("sheet1", "A1"), 3)
    pc.dump_stats(f'logs/test_reference_rename_sheet_{rows}.stats')
    with open(f'logs/test_reference_rename_sheet_stats_{rows}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_reference_rename_sheet_{rows}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()

def dangling_chain(self, width, height):
    pc = cProfile.Profile()
    wb = Workbook()
    wb.new_sheet("sheet1")
    wb.set_cell_contents("sheet1", "A1", "10")
    for y in range(1, width):
        c = chr(65 + y)
        c1 = chr(64 + y)
        wb.set_cell_contents("sheet1", f"{c}1", f"={c1}1")
    for y in range(width):
        c = chr(65 + y)
        for i in range(2, height):
            wb.set_cell_contents("sheet1", f"{c}{i}", f"={c}{i-1}")
    pc.enable()
    wb.set_cell_contents("sheet1", "A1", "20")
    pc.disable()
    self.assertEqual(wb.get_cell_value("sheet1", f"{chr(64+width)}{height-1}"), 20)
    pc.dump_stats(f'logs/test_dangling_chain_{height}_{width}.stats')
    with open(f'logs/test_dangling_chain_stats_{height}_{width}.stats', 'w') as stream:
        p = pstats.Stats(f'logs/test_dangling_chain_{height}_{width}.stats', stream=stream)
        p.sort_stats(pstats.SortKey.CUMULATIVE).print_stats()


class Multiple_Reference_Tests(unittest.TestCase):
    def test_many_reference_one(self):
        many_reference_one(self, 50)
        many_reference_one(self, 100)
        many_reference_one(self, 150)

    def test_many_reference_many(self):
        many_reference_many(self, 5, 10)
        many_reference_many(self, 10, 10)
        many_reference_many(self, 15, 10)

    def test_chain(self):
        chain(self, 50)
        chain(self, 100)
        chain(self, 150)
         
    def test_large_cycle(self):
        large_cycle(self, 50)
        large_cycle(self, 100)
        large_cycle(self, 150)

    def test_delete_sheet(self):
        delete_sheet1(self, 5, 10)
        delete_sheet1(self, 10, 10)
        delete_sheet1(self, 15, 10)

    def test_make_break_cycle(self):
        make_break_cycle(self, 20, 3)
        make_break_cycle(self, 40, 3)
        make_break_cycle(self, 60, 3)
        make_break_cycle(self, 20, 6)
        make_break_cycle(self, 20, 9)

    def test_rename_sheet(self):
        reference_rename_sheet(self, 5, 10)
        reference_rename_sheet(self, 10, 10)
        reference_rename_sheet(self, 15, 10)

    def test_dangling_chain(self):
        dangling_chain(self, 5, 10)
        dangling_chain(self, 10, 10)
        dangling_chain(self, 15, 10)


if __name__ == "__main__":
    unittest.main()