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
            p.sort_stats(pstats.SortKey.TIME).print_stats()
        cProfile.run('re.compile("foo|bar")')

class Multiple_Reference_Tests(unittest.TestCase):        
    def test_many_reference_one(self):
        pc = cProfile.Profile()
        pc.enable()
        wb = Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", "= 5000 * 2")
        for c in ["B", "C", "D", "E", "F", "G", "H", "I"]:
            for i in range(2, 100):
                wb.set_cell_contents("sheet1", f"{c}{i}", f"=A1")
        wb.set_cell_contents("sheet1", "A1", "= 5000")
        pc.disable()
        pc.dump_stats('logs/test_many_reference_one.stats')
        with open('logs/test_many_reference_one_stats.stats', 'w') as stream:
            p = pstats.Stats('logs/test_many_reference_one.stats', stream=stream)
            p.sort_stats(pstats.SortKey.TIME).print_stats()

        # for i in range(2, 1000):

        
        
        

        
        


if __name__ == "__main__":
    unittest.main()