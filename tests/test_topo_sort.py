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


class GitAction(unittest.TestCase):
    def test_lecture_example(self):
        wb = Workbook()
        wb.new_sheet("S1")
        # complete when set cell contents is implemented


# if __name__ == '__main__':
#     unittest.main()
