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
import sheets # noqa

class System_Tests(unittest.TestCase):
    def test_system1(self):
        self.assertEqual(sheets.version, "1.0")

        wb = sheets.Workbook()
        wb.new_sheet()
        


if __name__ == "__main__":
    unittest.main()