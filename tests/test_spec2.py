"""
Generic Integration tests adhering to Specification 1. 
"""

import unittest
import decimal
import json
from context import sheets


class Spec2_Tests(unittest.TestCase):
    """
    Tests adhering to specific examples presented in the Project 2
    specification.
    """

    def test_loading_workbook_basic(self):
        with open("test-data/mock_workbook.json", "r", encoding="utf8") as fp:
            wb = sheets.Workbook.load_workbook(fp)
            self.assertEqual(wb.get_cell_value("sheet2", "A1"),
                             decimal.Decimal('651.9'))
            self.assertEqual(wb.get_cell_value("sheet2", "A2"),
                             '"Hello, world"')
            with open("test-outputs/test_workbook2.json", "w",
                      encoding="utf8") as fpw:
                wb.save_workbook(fpw)

    def test_save_workbook_double_quotes(self):
        """
        JSON strings are always double-quoted; therefore,
        double-quotes within cell formulas must be escaped. Fortunately,
        the json library will take care of this for you -
        but you should still make sure it works properly. (visually)
        """
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "A1", '"hel"lo"')
        with open("test-outputs/test_save_workbook_double_quotes.json", "w",
                  encoding="utf8") as fpw:
            wb.save_workbook(fpw)
        with open("test-outputs/test_save_workbook_double_quotes.json", "r",
                  encoding="utf8") as fj:
            data = json.load(fj)
            self.assertEqual(data["sheets"][0]["cell-contents"]['A1'],
                             '"hel"lo"')
            # VISUAL CHECK

    def test_load_workbook_location_uppercase_lowercase(self):
        """
        When you read in a spreadsheet’s cells from the cell-contents JSON
        object, your code should support cell names being either uppercase
        or lowercase.
        """
        with open("test-data/mockwb_load_location_uppercase_lowercase.json",
                  "r", encoding="utf8") as fp:
            wb = sheets.Workbook.load_workbook(fp)
            self.assertEqual(wb.get_cell_value("sheet2", "A1"),
                             decimal.Decimal('651.9'))
            self.assertEqual(wb.get_cell_value("sheet2", "A2"),
                             "\"Hello, world\"")

    def test_save_workbook_location_uppercase(self):
        """
        Similarly, when you write out a spreadsheet’s cells, you may choose
        to output cell locations in either all-uppercase or all-lowercase,
        but pick one and be consistent."""
        wb = sheets.Workbook()
        wb.new_sheet("sheet1")
        wb.set_cell_contents("sheet1", "a1", "1")
        wb.set_cell_contents("sheet1", "A2", "2")
        wb.set_cell_contents("sheet1", "b1", "3")
        wb.set_cell_contents("sheet1", "b2", "4")
        with open("test-outputs/test_save_workbook_location_uppercase.json",
                  "w", encoding="utf8") as f:
            wb.save_workbook(f)
        with open("test-outputs/test_save_workbook_location_uppercase.json",
                  "r", encoding="utf8") as fj:
            data = json.load(fj)
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
        with open("test-data/mockwb_bad_json1.json", "r",
                  encoding="utf8") as f:
            with self.assertRaises(KeyError):
                _ = sheets.Workbook.load_workbook(f)

    def test_bad_json2(self):
        with open("test-data/mockwb_bad_json2.json", "r",
                  encoding="utf8") as f:
            with self.assertRaises(KeyError):
                _ = sheets.Workbook.load_workbook(f)

    def test_bad_json3(self):
        with open("test-data/mockwb_bad_json3.json", "r",
                  encoding="utf8") as f:
            with self.assertRaises(KeyError):
                _ = sheets.Workbook.load_workbook(f)

    def test_bad_json4(self):
        with open("test-data/mockwb_bad_json4.json", "r",
                  encoding="utf8") as f:
            with self.assertRaises(TypeError):
                _ = sheets.Workbook.load_workbook(f)

    def test_bad_json_bad_reference(self):
        with open("test-data/mockwb_bad_json_bad_reference.json", "r",
                  encoding="utf8") as f:
            wb = sheets.Workbook.load_workbook(f)
            value = wb.get_cell_value('sheet2', 'A1')
            self.assertTrue(isinstance(value, sheets.CellError))
            self.assertTrue(value.get_type() ==
                            sheets.CellErrorType.BAD_REFERENCE)

    def test_bad_json_parse_error(self):
        with open("test-data/mockwb_bad_json_parse_error.json", "r",
                  encoding="utf8") as f:
            wb = sheets.Workbook.load_workbook(f)
            value = wb.get_cell_value('sheet1', 'C1')
            self.assertTrue(isinstance(value, sheets.CellError))
            self.assertTrue(value.get_type() ==
                            sheets.CellErrorType.PARSE_ERROR)

    def test_copy_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.copy_sheet('sheet1')
        self.assertEqual(list(wb.spreadsheets.keys())[1], 'sheet1_1')
        wb.copy_sheet('sheet1')
        self.assertEqual(list(wb.spreadsheets.keys())[2], 'sheet1_2')
        wb.copy_sheet('sheet1_1')
        self.assertEqual(list(wb.spreadsheets.keys())[3], 'sheet1_1_1')


if __name__ == "__main__":
    unittest.main()
