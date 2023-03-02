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

import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

class TestSCC(unittest.TestCase):
    def test_scc(self):
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A2", "=A11")
        wb.set_cell_contents("sheet1", "B9", "=B8 + A11")
        wb.set_cell_contents("sheet1", "A11", "=A5 + B7")
        wb.set_cell_contents("sheet1", "B8", "=B7 + C3")
        wb.set_cell_contents("sheet1", "C10", "=C3 + A11")
        wb.set_cell_contents("sheet1", "A5", "")
        wb.set_cell_contents("sheet1", "B7", "")
        wb.set_cell_contents("sheet1", "C3", "")

        # wb = sheets.Workbook()
        # wb.new_sheet("sheet1")
        # wb.set_cell_contents("sheet1", "A1", "=A2 + B1")
        # wb.set_cell_contents("sheet1", "B1", "=A1 + C1")
        # wb.set_cell_contents("sheet1", "C1", "")
        # wb.set_cell_contents("sheet1", "A2", "")
        
        # wb = sheets.Workbook()
        # wb.new_sheet()
        # wb.set_cell_contents('sheet1', 'A1', '=B1 + D1')
        # wb.set_cell_contents('sheet1', 'B1', '=C1')
        # wb.set_cell_contents('sheet1', 'D1', '=C1')
        # wb.set_cell_contents('sheet1', 'C1', '10')

        # wb = sheets.Workbook()
        # wb.new_sheet()
        # wb.set_cell_contents('sheet1', 'C1', '5')
        # wb.set_cell_contents('sheet1', 'B1', '=A1')
        # wb.set_cell_contents('sheet1', 'A1', '=if(a2,b1,c1)')
        # wb.set_cell_contents('sheet1', 'A2', 'TRUE')

        s = sheets.tarjan(wb.spreadsheets["sheet1"].cells["A2"], wb.adjacency_list)
        s = s[::-1]
        for l in s:
            lst = [i.location for i in l]

if __name__ == "__main__":
    unittest.main()