from typing import *
from sheets.lark_module import FormulaEvaluator
from . import Sheet
from sheets import cell
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
        self.lower_names = set()  # set of lower case name
        self.curr_lowest = 1  # current open

    def num_sheets(self) -> int:
        return len(self.spreadsheets)

    def get_cell(self, sheet_name: str, location: str) -> cell.Cell:
        return self.spreadsheets[sheet_name].cells[location]

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
            i = self.curr_lowest
            while True:
                sheet_name = "Sheet" + str(i)
                if sheet_name.lower() not in self.lower_names:
                    self.spreadsheets[sheet_name.lower()] = Sheet(sheet_name)
                    self.lower_names.add(sheet_name.lower())
                    self.curr_lowest = i + 1
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
        elif sheet_name.lower() in self.lower_names:
            raise ValueError("Duplicate spreadsheet name")
        elif sheet_name.lower() not in self.lower_names:
            self.spreadsheets[sheet_name.lower()] = Sheet(sheet_name)
            self.lower_names.add(sheet_name.lower())
            return (len(self.spreadsheets) - 1, sheet_name)

    def __update_values(self, cell):
        contents = cell.contents
        _, value = lark_module.evaluate_expr(self, cell, cell.sheet, contents)
        cell.value = value

    def __update_extent(self, sheet, location, deletingCell: bool):
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

    def __is_number(self, string):
        return string.isnumeric() or (string.replace('.', '', 1).isdigit() and string.count('.') < 2)

    def __strip_zeros(self, string):
        return string.rstrip('0').rstrip('.') if '.' in string else string

    def __strip_evaluation(self, eval):
        # given evaluation from lark, return value
        if type(eval) == str and self.__is_number(eval):
            contents = self.__strip_zeros(eval)
            return decimal.Decimal(contents)
        elif type(eval) == decimal.Decimal:
            stripped = self.__strip_zeros(str(eval))
            return decimal.Decimal(stripped)
        else:
            return eval

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.lower_names:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        del self.spreadsheets[sheet_name.lower()]
        for loc in sheet.cells:
            c = sheet.cells[loc]
            past_relies_on = c.relies_on
            list_diff = past_relies_on - c.relies_on
            for c in list_diff:
                c.dependents.remove(c)
            # update dependents
            _, sorted_components = topo_sort(c)
            for node in sorted_components[1:]:
                self.__update_values(node)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.lower_names:
            raise KeyError("Specified sheet name not found")
        else:
            sheet = self.spreadsheets[sheet_name.lower()]
            return ((sheet.extent_col, sheet.extent_row))

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
        if sheet_name.lower() not in self.lower_names:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        if not sheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if contents:
            contents = contents.strip()
        if location in sheet.cells:
            curr_cell = sheet.cells[location]
            past_relies_on = curr_cell.relies_on
            # No one depends on this cell: delete without traversal
            if (not contents or len(contents) == 0) and len(curr_cell.dependents) == 0:
                del sheet.cells[location]
                self.__update_extent(sheet, location, True)
                return
            # Some cell depends on this cell: delete with traversal
            elif not contents or len(contents) == 0:
                curr_cell.set_fields(contents=None, value=None, relies_on=set(),
                                     type=cell.CellType.EMPTY)
                circular, sorted_components = topo_sort(curr_cell)
                if not circular:
                    for node in sorted_components[1:]:
                        self.__update_values(node)
                else:
                    for node in curr_cell.dependents:
                        node.relies_on.remove(curr_cell)
                    for node in sorted_components[1:]:
                        self.__update_values(node)
                self.__update_extent(sheet, location, True)
                return
            curr_cell.contents = contents
            # Update cell contents and value
            if contents[0] == "=":
                eval, value = lark_module.evaluate_expr(
                    self, curr_cell, sheet.name, contents)
                value = self.__strip_evaluation(value)
                curr_cell.set_fields(
                    value=value, type=cell.CellType.FORMULA, relies_on=eval.relies_on)
            elif contents[0] == "'":
                curr_cell.set_fields(
                    value=contents[1:].rstrip(), type=cell.CellType.STRING, relies_on=set())
            elif self.__is_number(contents):
                contents = self.__strip_zeros(contents)
                curr_cell.set_fields(value=decimal.Decimal(
                    contents), type=cell.CellType.LITERAL_NUM, relies_on=set())
            else:
                curr_cell.set_fields(
                    value=contents, type=cell.CellType.LITERAL_STRING, relies_on=set())
            # update relies on
            list_diff = past_relies_on - curr_cell.relies_on
            for c in list_diff:
                c.dependents.remove(curr_cell)
            # update dependents
            circular, sorted_components = topo_sort(curr_cell)
            if not circular:
                for node in sorted_components[1:]:
                    self.__update_values(node)
            else:
                for node in sorted_components:
                    node.value = cell_error.CellError(
                        cell_error.CellErrorType.CIRCULAR_REFERENCE, 'cycle detected')

        # cell does not exist: add cell
        else:
            if not contents or len(contents) == 0:
                return
            if contents[0] == "=":
                curr_cell = cell.Cell(
                    sheet_name, location, contents, None, cell.CellType.FORMULA)
                eval, value = lark_module.evaluate_expr(
                    self, curr_cell, sheet_name, contents)
                # FIX THIS
                value = self.__strip_evaluation(value)
                curr_cell.set_fields(value=value, relies_on=eval.relies_on)
                sheet = self.spreadsheets[sheet_name.lower()]
                sheet.cells[location] = curr_cell
            elif contents[0] == "'":
                value = contents[1:].rstrip()
                curr_cell = cell.Cell(
                    sheet_name, location, contents, value, cell.CellType.STRING)
                sheet = self.spreadsheets[sheet_name.lower()]
                sheet.cells[location] = curr_cell
            elif self.__is_number(contents):
                contents = self.__strip_zeros(contents)
                value = decimal.Decimal(contents)
                curr_cell = cell.Cell(
                    sheet_name, location, contents, value, cell.CellType.LITERAL_NUM)
                sheet = self.spreadsheets[sheet_name.lower()]
                sheet.cells[location] = curr_cell
            else:
                value = contents
                curr_cell = cell.Cell(
                    sheet_name, location, contents, value, cell.CellType.LITERAL_STRING)
                sheet = self.spreadsheets[sheet_name.lower()]
                sheet.cells[location] = curr_cell
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
        if sheet_name.lower() not in self.lower_names:
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
        if sheet_name.lower() not in self.lower_names:
            raise KeyError("Specified sheet name not found")
        sheet = self.spreadsheets[sheet_name.lower()]
        if not sheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if location in sheet.cells:
            return sheet.cells[location].value
        else:
            return None
