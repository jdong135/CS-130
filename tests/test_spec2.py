"""
Generic Integration tests adhering to Specification 1. 
"""

import unittest
import os
import sys
import decimal
import json

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import Workbook, lark_module, CellErrorType, CellError  # noqa

class Spec2_Tests(unittest.TestCase):
    def test_loading_workbook_basic(self):
        fp = open("test-data/mock_workbook.json", "r")
        wb = Workbook.load_workbook(fp)
        fp.close()
        self.assertEqual(wb.get_cell_value("sheet2", "A1"),
                         decimal.Decimal('651.9'))
        self.assertEqual(wb.get_cell_value("sheet2", "A2"), '"Hello, world"')
        fpw = open("test-outputs/test_workbook2.json", "w")
        wb.save_workbook(fpw)
        fpw.close()

    def test_save_workbook_double_quotes(self):
        """
        JSON strings are always double-quoted; therefore,
        double-quotes within cell formulas must be escaped. Fortunately,
        the json library will take care of this for you -
        but you should still make sure it works properly. (visually)
        """
        wb = Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", '"hel"lo"')
        fpw = open("test-outputs/test_save_workbook_double_quotes.json", "w")
        wb.save_workbook(fpw)
        fpw.close()
        fj = open("test-outputs/test_save_workbook_double_quotes.json", "r")
        data = json.load(fj)
        self.assertEqual(data["sheets"][0]["cell-contents"]['A1'], '"hel"lo"')
        fj.close()
        # VISUAL CHECK

    def test_load_workbook_location_uppercase_lowercase(self):
        """
        When you read in a spreadsheet’s cells from the cell-contents JSON
        object, your code should support cell names being either uppercase
        or lowercase.
        """
        fp = open("test-data/mockwb_load_location_uppercase_lowercase.json", "r")
        wb = Workbook.load_workbook(fp)
        fp.close()
        self.assertEqual(wb.get_cell_value("sheet2", "A1"),
                         decimal.Decimal('651.9'))
        self.assertEqual(wb.get_cell_value("sheet2", "A2"), "\"Hello, world\"")

    def test_save_workbook_location_uppercase(self):
        """
        Similarly, when you write out a spreadsheet’s cells, you may choose
        to output cell locations in either all-uppercase or all-lowercase,
        but pick one and be consistent."""
        wb = Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "a1", "1")
        wb.set_cell_contents("sheet1", "A2", "2")
        wb.set_cell_contents("sheet1", "b1", "3")
        wb.set_cell_contents("sheet1", "b2", "4")
        f = open("test-outputs/test_save_workbook_location_uppercase.json", "w")
        wb.save_workbook(f)
        f.close()
        fj = open("test-outputs/test_save_workbook_location_uppercase.json", "r")
        data = json.load(fj)
        fj.close()
        self.assertEqual(data["sheets"][0]["cell-contents"]['A1'], '1')
        self.assertEqual(data["sheets"][0]["cell-contents"]['A2'], '2')
        self.assertEqual(data["sheets"][0]["cell-contents"]['B1'], '3')
        self.assertEqual(data["sheets"][0]["cell-contents"]['B2'], '4')
        # VISUAL CHECK
    
    def test_bad_json1(self):
        """
        Since we are focusing on implementing high-quality software, the code
        that loads a workbook from a JSON file will need to gracefully detect
        and report bad inputs; for example, malformed/unparseable JSON, valid
        JSON data that is missing the required fields or has the wrong types,
        invalid cell references, etc.
        """
        f = open("test-data/mockwb_bad_json1.json", "r")
        with self.assertRaises(KeyError):
            wb = Workbook.load_workbook(f)
        f.close()
    
    def test_bad_json2(self):
        f = open("test-data/mockwb_bad_json2.json", "r")
        with self.assertRaises(KeyError):
            wb = Workbook.load_workbook(f)
        f.close()

    def test_bad_json3(self):
        f = open("test-data/mockwb_bad_json3.json", "r")
        with self.assertRaises(KeyError):
            wb = Workbook.load_workbook(f)
        f.close()

    def test_bad_json4(self):
        f = open("test-data/mockwb_bad_json4.json", "r")
        with self.assertRaises(TypeError):
            wb = Workbook.load_workbook(f)
        f.close()

    def test_bad_json_bad_reference(self):
        f = open("test-data/mockwb_bad_json_bad_reference.json", "r")
        wb = Workbook.load_workbook(f)
        f.close()
        value = wb.get_cell_value('sheet2', 'A1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

    def test_bad_json_parse_error(self):
        f = open("test-data/mockwb_bad_json_parse_error.json", "r")
        wb = Workbook.load_workbook(f)
        f.close()
        value = wb.get_cell_value('sheet1', 'C1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR


if __name__ == "__main__":
    unittest.main()
