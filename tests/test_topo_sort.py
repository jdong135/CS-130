"""
Test for implementation of Topological Sort algorithm. 
"""

import unittest
import os
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import *  # noqa


class Topo_Sort(unittest.TestCase):
    def test_two(self):
        wb = Workbook()
        wb.new_sheet("S1")
        # wb.set_cell_contents("S1", "A1", "=4")
        # wb.set_cell_contents("S1", "A2", "=A1")
        # c1 = wb.get_cell("S1", "A1")
        # sort = topo_sort(c1)
        # print(sort)


if __name__ == '__main__':
    unittest.main()
