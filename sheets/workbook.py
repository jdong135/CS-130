from typing import *
from sheets.lark_module import FormulaEvaluator
from . import Sheet
from sheets import cell, strip_module
from sheets import topo_sort
from sheets import cell_error
import lark
import decimal
from sheets import lark_module


ALLOWED_PUNC = set([".", "?", "!", ",", ":", ";", "@", "#",
                    "$", "%", "^", "&", "*", "(", ")", "-", "_"])


class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        self.spreadsheets = {}  # lower case name -> sheet object
        self.adjacency_list = {}  # Cell: [neighbor Cells]

    def num_sheets(self) -> int:
        """
        Return current number of spreadsheets
        """
        return len(self.spreadsheets)

    def list_sheets(self) -> List[str]:
        # Return a list of the spreadsheet names in the workbook, with the
        # capitalization specified at creation, and in the order that the sheets
        # appear within the workbook.
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.
        output = []
        for sheet in self.spreadsheets:
            output.append(self.spreadsheets[sheet].name)
        return output

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        # Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated.  "Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.
        if sheet_name == "":
            raise ValueError("Sheet name is empty string")
        elif not sheet_name:
            # handle null name
            i = 0
            while True:
                sheet_name = "Sheet" + str(i + 1)
                if sheet_name.lower() not in self.spreadsheets:
                    self.spreadsheets[sheet_name.lower()] = Sheet(sheet_name)
                    return (len(self.spreadsheets) - 1, sheet_name)
                i += 1
        for ch in sheet_name:
            if not ch.isalnum() and ch != " " and ch not in ALLOWED_PUNC:
                raise ValueError(
                    f"Invalid character in name: {ch} not allowed")
        if sheet_name[0] == " ":
            raise ValueError("Sheet name starts with white space")
        elif sheet_name[-1] == " ":
            raise ValueError("Sheet name ends with white space")
        elif sheet_name.lower() in self.spreadsheets:
            raise ValueError("Duplicate spreadsheet name")
        elif sheet_name.lower() not in self.spreadsheets:
            self.spreadsheets[sheet_name.lower()] = Sheet(sheet_name)
            return (len(self.spreadsheets) - 1, sheet_name)

    def __update_extent(self, sheet, location, deletingCell: bool):
        """
        Update the extent of a sheet if we are deleting a cell
        """
        if deletingCell:
            sheet_col, sheet_row = sheet.extent_col, sheet.extent_row
            col, row = sheet.str_to_tuple(location)
            max_col, max_row = 0, 0
            if col == sheet_col or row == sheet_row:
                for c in sheet.cells:
                    if sheet.cells[c].type != cell.CellType.EMPTY:
                        c_col, c_row = sheet.str_to_tuple(
                            sheet.cells[c].location)
                        max_col = max(max_col, c_col)
                        max_row = max(max_row, c_row)
                sheet.extent_col = max_col
                sheet.extent_row = max_row
        else:
            curr_col, curr_row = sheet.str_to_tuple(location)
            sheet.extent_col = max(curr_col, sheet.extent_col)
            sheet.extent_row = max(curr_row, sheet.extent_row)

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        for c in sheet.cells:
            for _, neighbors in self.adjacency_list.items():
                if c in neighbors:
                    neighbors.remove(c)
            cell_dependents = topo_sort(c, self.adjacency_list)[1:]
            for dependent in cell_dependents:
                self.__set_cell_value_and_type(dependent)
            del self.adjacency_list[c]
        del self.spreadsheets[sheet_name.lower()]

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError("Specified sheet name not found")
        else:
            sheet = self.spreadsheets[sheet_name.lower()]
            return ((sheet.extent_col, sheet.extent_row))

    def __str_to_error(self, str_error: str):
        if str_error == "#ERROR!":
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "input error")
        elif str_error == "#CIRCREF!":
            return cell_error.CellError(cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
        elif str_error == "#REF!":
            return cell_error.CellError(cell_error.CellErrorType.BAD_REFERENCE, "input error")
        elif str_error == "#NAME?":
            return cell_error.CellError(cell_error.CellErrorType.BAD_NAME, "input error")
        elif str_error == "#VALUE!":
            return cell_error.CellError(cell_error.CellErrorType.TYPE_ERROR, "input error")
        elif str_error == "#DIV/0!":
            return cell_error.CellError(cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")
        else:
            return None

    def __set_cell_value_and_type(self, calling_cell):
        """
        Sets cells value and type based on cell's contents field
        Args:
            cell (_type_): Assume cell has contents set correctly
        """
        cell_contents = calling_cell.contents
        if not cell_contents or len(cell_contents) == 0:
            val = None
            type = cell.CellType.EMPTY
        elif self.__str_to_error(cell_contents):  # type error
            val = self.__str_to_error(cell_contents)
            type = cell.CellType.ERROR
        elif cell_contents[0] == "'":
            val = cell_contents[1:]
            type = cell.CellType.STRING
        elif cell_contents[0] == "=":
            val = lark_module.evaluate_expr(
                self, calling_cell, calling_cell.sheet, cell_contents)
            type = cell.CellType.FORMULA
        elif strip_module.is_number(cell_contents):
            val = decimal.Decimal(cell_contents)
            type = cell.CellType.LITERAL_NUM
        else:
            val = cell_contents
            type = cell.CellType.LITERAL_STRING
        calling_cell.set_fields(value=val, type=type)

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        # Set the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # naure of the issue.

        # Set cell contents and then evaluate with lark
        location = location.upper()
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        if not sheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if contents:
            contents = contents.strip()
        if location in sheet.cells:  # if cell already exists (modify contents)
            existing_cell = sheet.cells[location]
            existing_cell.contents = contents
            self.__set_cell_value_and_type(existing_cell)
            # if no cells depend on existing cell and contents are empty, delete cell
            if existing_cell.type == cell.CellType.EMPTY:
                if not self.adjacency_list[existing_cell]:
                    del sheet.cells[location]
                    del self.adjacency_list[existing_cell]
                    self.__update_extent(sheet, location, True)
                    for _, neighbors in self.adjacency_list.items():
                        if existing_cell in neighbors:
                            neighbors.remove(existing_cell)
                    return
            # updating cells that depend on existing cell
            cell_dependents = topo_sort(existing_cell, self.adjacency_list)[1:]
            for dependent in cell_dependents:
                self.__set_cell_value_and_type(dependent)
        else:  # if cell does not exist (create contents)
            new_cell = cell.Cell(sheet_name, location, contents, None, None)
            self.__set_cell_value_and_type(new_cell)
            if new_cell.type != cell.CellType.EMPTY:
                self.adjacency_list[new_cell] = []
                sheet.cells[location] = new_cell
            self.__update_extent(sheet, location, False)

    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        # Return the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.
        location = location.upper()
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        if not sheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if location in sheet.cells:
            return sheet.cells[location].contents
        else:
            return None

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        # Return the evaluated value of the specified cell on the specified
        # sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').

        # Return evaluation field of cell
        location = location.upper()
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        if not sheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if location in sheet.cells:
            return sheet.cells[location].value
        else:
            return None
