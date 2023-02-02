from __future__ import annotations
import logging
from typing import *
from sheets import cell, topo_sort, cell_error, lark_module, sheet, string_conversions
import decimal
import copy
import json
import re

ALLOWED_PUNC = set([".", "?", "!", ",", ":", ";", "@", "#",
                    "$", "%", "^", "&", "*", "(", ")", "-", "_"])

logging.basicConfig(filename="logs/results.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        # lower case name -> sheet object
        self.spreadsheets = {}
        # Cell: [neighbor Cells]; neighbors are cells that depend on Cell
        self.adjacency_list = {}
        # notify functions = set of user-inputted notify functions
        self.notify_functions = []

    def __check_valid_sheet_name(self, sheet_name: str):
        """
        Check if a sheet name is valid. 

        Args:
            sheet_name (str): Sheet name to check

        Raises:
            ValueError: Invalid character is used
            ValueError: Sheet name starts with white space
            ValueError: Sheet name ends with white space
            ValueError: Duplicate sheet name
        """
        if len(sheet_name) < 1 or not sheet_name:
            raise ValueError("Empty string for sheet name.")
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

    def __update_extent(self, sheet:sheet.Sheet , location: str, deletingCell: bool):
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

    def __generate_notifications(self, cell_list: list[cell.Cell]):
        """
        Given a list of cells, create a corresponding list of tuples containing
        each cell's sheet name and location. Then, call each registered notify
        function on the created list. 

        Args:
            cell_list (list): List of Cell objects to generate a list of 
            notifications from. 
        """
        if cell_list == []:
            return
        changed_cells = []
        for c in cell_list:
            changed_cells.append((c.sheet.name, c.location))
        for notify_function in self.notify_functions:
            try:
                notify_function(copy.deepcopy(self), changed_cells)
            except:
                continue

    def __set_cell_value_and_type(self, calling_cell: cell.Cell) -> Tuple[list, bool]:
        """
        Sets cells value and type based on cell's contents field

        Args:
            calling_cell (Cell): Cell object that we assume has contents set correctly.

        Returns:
            list: cells that the calling cell relies on.
            bool: if the value of the calling cell changed
        """
        cell_contents = calling_cell.contents
        relies_on = []
        if not cell_contents or len(cell_contents) == 0:
            val = None
            type = cell.CellType.EMPTY
        elif string_conversions.str_to_error(cell_contents):  # type error
            val = string_conversions.str_to_error(cell_contents)
            type = cell.CellType.ERROR
        elif cell_contents[0] == "'":
            val = cell_contents[1:]
            type = cell.CellType.STRING
        elif cell_contents[0] == "=":
            eval, val = lark_module.evaluate_expr(
                self, calling_cell, calling_cell.sheet.name, cell_contents)
            type = cell.CellType.FORMULA
            if eval:
                relies_on = eval.calling_cell_relies_on
        elif string_conversions.is_number(cell_contents):
            stripped = string_conversions.strip_zeros(cell_contents)
            val = decimal.Decimal(stripped)
            type = cell.CellType.LITERAL_NUM
        else:
            val = cell_contents
            type = cell.CellType.LITERAL_STRING
        val_update = True if val != calling_cell.value or type != calling_cell.type else False
        calling_cell.set_fields(value=val, type=type)
        return relies_on, val_update

    def __get_cells_containing_sheetname(self, sheetname: str) -> list[cell.Cell]:
        # match any cell that has contents sheetname! or 'sheetname'!
        cells = []
        regex = f'.*{sheetname.lower()}!.*'
        regex2 = f".*'{sheetname.lower()}'!.*"
        for name, sheet in self.spreadsheets.items():
            for loc, c in sheet.cells.items():
                if c.contents and (re.match(regex, c.contents.lower()) or re.match(regex2, c.contents.lower())):
                    try:
                        # If a cell is an error type, check if it is a parse error and don't
                        # add to our list of cells. Otherwise, if it is not a cell error type,
                        # get_type() will throw an attribute error.
                        if self.get_cell_value(name, loc).get_type() == cell_error.CellErrorType.PARSE_ERROR:
                            continue
                    except AttributeError:
                        pass
                    cells.append(c)
        return cells

    def __get_cell_contents_after_rename(self, c: cell.Cell , sheet_name: str, new_sheet_name: str):
        # ensure names with spaces are wrapped in quotes
        if ' ' in new_sheet_name:
            new_sheet_name = f"'{new_sheet_name}'"
        new_contents = re.sub(f"{sheet_name}!|'{sheet_name}'!",
                              f"{new_sheet_name}!", c.contents, flags=re.IGNORECASE)
        # remove '' around names that don't need quotations
        quoted = re.findall(r"'[\w\s.?!,:;!@#$%^&*()-_]+'!", new_contents)
        for name in quoted:
            if ' ' not in name:
                # name = 'sheetname'!
                new_contents = re.sub(
                    f"{name}", f"{name[1:-2]}!", new_contents, flags=re.IGNORECASE)
        return new_contents

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
        elif sheet_name:
            self.__check_valid_sheet_name(sheet_name)
        else: # handle null input 
            i = 1
            while True:
                sheet_name = "Sheet" + str(i)
                if sheet_name.lower() not in self.spreadsheets:
                    break
                i += 1
        self.spreadsheets[sheet_name.lower()] = sheet.Sheet(sheet_name)
        changed_cells = []
        for c in list(self.adjacency_list.keys()):
            _, c_updated = self.__set_cell_value_and_type(c)
            if c_updated:
                changed_cells.append(c)
        self.__generate_notifications(changed_cells)
        return len(self.spreadsheets) - 1, sheet_name

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
        del self.spreadsheets[sheet_name.lower()]
        changed_cells = []
        for loc in sheet.cells:
            c = sheet.cells[loc]
            _, cell_dependents = topo_sort(c, self.adjacency_list)
            changed_cells.extend(cell_dependents[1:])
            for _, neighbors in self.adjacency_list.items():
                if c in neighbors:
                    neighbors.remove(c)
            for dependent in cell_dependents[1:]:
                self.__set_cell_value_and_type(dependent)
            del self.adjacency_list[c]
        self.__generate_notifications(changed_cells)

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
            relies_on, val_updated = self.__set_cell_value_and_type(existing_cell)
            # Everything the existing cell relies on
            for c, neighbors in self.adjacency_list.items():
                if existing_cell in neighbors and c not in relies_on:
                    neighbors.remove(existing_cell)       
            if existing_cell.type == cell.CellType.EMPTY:
                # existing cell is now empty so it does not depend on other cells
                # -> remove it as a neighbor of other cells
                for _, neighbors in self.adjacency_list.items():
                    if existing_cell in neighbors:
                        neighbors.remove(existing_cell)
                # if existing cell doesn't have neighbors, no cell relies on it
                # -> delete cell from spreadsheet
                if not self.adjacency_list[existing_cell]:
                    del sheet.cells[location]
                    del self.adjacency_list[existing_cell]
                    self.__update_extent(sheet, location, True)
                    self.__generate_notifications([existing_cell])
                    return
            # updating cells that depend on existing cell
            circular, cell_dependents = topo_sort(
                existing_cell, self.adjacency_list)
            if not circular:
                for dependent in cell_dependents[1:]:
                    self.__set_cell_value_and_type(dependent)
            else:  # everything in the cycle should have an error value
                for dependent in cell_dependents:
                    dependent.set_fields(value=cell_error.CellError(
                        cell_error.CellErrorType.CIRCULAR_REFERENCE, "circular reference"))
            self.__update_extent(sheet, location, False)
            # include the existing cell iff its value is updated
            self.__generate_notifications(cell_dependents if val_updated else cell_dependents[1:])
        else:  # if cell does not exist (create contents)
            new_cell = cell.Cell(sheet, location, contents, None, None)
            self.__set_cell_value_and_type(new_cell)
            if new_cell.type != cell.CellType.EMPTY:
                self.adjacency_list[new_cell] = []
                sheet.cells[location] = new_cell
            self.__update_extent(sheet, location, False)
            self.__generate_notifications([new_cell])

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

    @staticmethod
    def load_workbook(fp: TextIO) -> Workbook:
        # This is a static method (not an instance method) to load a workbook
        # from a text file or file-like object in JSON format, and return the
        # new Workbook instance.  Note that the _caller_ of this function is
        # expected to have opened the file; this function merely reads the file.
        #
        # If the contents of the input cannot be parsed by the Python json
        # module then a json.JSONDecodeError should be raised by the method.
        # (Just let the json module's exceptions propagate through.)  Similarly,
        # if an IO read error occurs (unlikely but possible), let any raised
        # exception propagate through.
        #
        # If any expected value in the input JSON is missing (e.g. a sheet
        # object doesn't have the "cell-contents" key), raise a KeyError with
        # a suitably descriptive message.
        #
        # If any expected value in the input JSON is not of the proper type
        # (e.g. an object instead of a list, or a number instead of a string),
        # raise a TypeError with a suitably descriptive message.
        wb = Workbook()
        try:
            data = json.load(fp)
        except json.JSONDecodeError as e:
            print(f"{e}, Decode Error in load_workbook().")
            return wb
        except IOError as e:
            print(f"{e}, IO Error in load_workbook().")
            return wb

        if len(data) != 1:
            raise TypeError("Should contain one instance of sheets.")
        if list(data.keys())[0] != "sheets":
            raise KeyError("Key should be named \"sheets\".")
        sheets = data["sheets"]
        if type(sheets) != list:
            raise TypeError("\"sheets\" value should be of type list")

        for sheet in sheets:
            if type(sheet) != dict:
                raise TypeError("Sheet object is not of type dictionary.")
            if len(sheet.keys()) != 2:
                raise KeyError(
                    "Sheet dictionary does not have exactly two keys.")

            if "name" not in sheet.keys() or "cell-contents" not in sheet.keys():
                raise KeyError(
                    "Keys of sheet dictionray must be named \"name\" and \"cell-contents\"")
            sheet_name = sheet["name"]
            if type(sheet_name) != str:
                raise TypeError("Sheet name is not of type string.")
            wb.new_sheet(sheet_name)

            cell_contents = sheet["cell-contents"]
            if type(cell_contents) != dict:
                raise TypeError(
                    f"Cell Contents of sheet \"{sheet_name}\" is not of type dict.")
            for location, contents in cell_contents.items():
                if type(location) != str or type(contents) != str:
                    raise TypeError(
                        "Cell location and contents must be of type string.")
                wb.set_cell_contents(sheet_name, location, contents)
        return wb

    def save_workbook(self, fp: TextIO) -> None:
        # Instance method (not a static/class method) to save a workbook to a
        # text file or file-like object in JSON format.  Note that the _caller_
        # of this function is expected to have opened the file; this function
        # merely writes the file.
        #
        # If an IO write error occurs (unlikely but possible), let any raised
        # exception propagate through.
        data = {"sheets": []}
        for _, sheet in self.spreadsheets.items():
            name = sheet.name
            cur_sheet = {'name': name, 'cell-contents': {}}
            for location, c in sheet.cells.items():
                cur_sheet['cell-contents'][location] = c.contents
            data["sheets"].append(cur_sheet)
        try:
            json.dump(data, fp, indent=4)
        except IOError as e:
            print(f"{e}, IO write error in save_workbook().")

    def notify_cells_changed(self,
                             notify_function: Callable[[Workbook, Iterable[Tuple[str, str]]], None]) -> None:
        # Request that all changes to cell values in the workbook are reported
        # to the specified notify_function.  The values passed to the notify
        # function are the workbook, and an iterable of 2-tuples of strings,
        # of the form ([sheet name], [cell location]).  The notify_function is
        # expected not to return any value; any return-value will be ignored.
        #
        # Multiple notification functions may be registered on the workbook;
        # functions will be called in the order that they are registered.
        #
        # A given notification function may be registered more than once; it
        # will receive each notification as many times as it was registered.
        #
        # If the notify_function raises an exception while handling a
        # notification, this will not affect workbook calculation updates or
        # calls to other notification functions.
        #
        # A notification function is expected to not mutate the workbook or
        # iterable that it is passed to it.  If a notification function violates
        # this requirement, the behavior is undefined.
        if not callable(notify_function):
            raise TypeError(
                f"Input notify function {notify_function} is not callable.")
        self.notify_functions.append(notify_function)

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        # Rename the specified sheet to the new sheet name.  Additionally, all
        # cell formulas that referenced the original sheet name are updated to
        # reference the new sheet name (using the same case as the new sheet
        # name, and single-quotes iff [if and only if] necessary).
        #
        # The sheet_name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # As with new_sheet(), the case of the new_sheet_name is preserved by
        # the workbook.
        #
        # If the sheet_name is not found, a KeyError is raised.
        #
        # If the new_sheet_name is an empty string or is otherwise invalid, a
        # ValueError is raised.
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError(f"Sheet name \"{sheet_name}\" not found.")
        self.__check_valid_sheet_name(new_sheet_name)
        # Get current index of our original sheet
        index = list(self.spreadsheets.keys()).index(sheet_name.lower())

        # Rename the current sheet
        self.spreadsheets[new_sheet_name.lower()] = self.spreadsheets.pop(
            sheet_name.lower())
        self.spreadsheets[new_sheet_name.lower()].name = new_sheet_name
        self.move_sheet(new_sheet_name, index)

        cells_to_update = self.__get_cells_containing_sheetname(sheet_name)
        for c in cells_to_update:
            new_contents = self.__get_cell_contents_after_rename(
                c, sheet_name, new_sheet_name)
            self.set_cell_contents(c.sheet.name, c.location, new_contents)

        ref_error_cells = self.__get_cells_containing_sheetname(new_sheet_name)
        for c in ref_error_cells:
            if c not in cells_to_update:
                self.set_cell_contents(c.sheet.name, c.location, c.contents)

    def move_sheet(self, sheet_name: str, index: int) -> None:
        # Move the specified sheet to the specified index in the workbook's
        # ordered sequence of sheets.  The index can range from 0 to
        # workbook.num_sheets() - 1.  The index is interpreted as if the
        # specified sheet were removed from the list of sheets, and then
        # re-inserted at the specified index.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        #
        # If the index is outside the valid range, an IndexError is raised.
        if index > self.num_sheets() - 1 or index < 0:
            raise IndexError("Index out of range")
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError(f"Sheet '{sheet_name}' not found")

        sheet_to_insert = self.spreadsheets[sheet_name.lower()]
        del self.spreadsheets[sheet_name.lower()]
        new_spreadsheets = {}
        for i, name in enumerate(self.spreadsheets):
            if i == index:
                break
            new_spreadsheets[name] = self.spreadsheets[name]
        new_spreadsheets[sheet_name.lower()] = sheet_to_insert
        for i, name in enumerate(self.spreadsheets):
            if i >= index:
                new_spreadsheets[name] = self.spreadsheets[name]
        self.spreadsheets = new_spreadsheets

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        # Make a copy of the specified sheet, storing the copy at the end of the
        # workbook's sequence of sheets.  The copy's name is generated by
        # appending "_1", "_2", ... to the original sheet's name (preserving the
        # original sheet name's case), incrementing the number until a unique
        # name is found.  As usual, "uniqueness" is determined in a
        # case-insensitive manner.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The copy should be added to the end of the sequence of sheets in the
        # workbook.  Like new_sheet(), this function returns a tuple with two
        # elements:  (0-based index of copy in workbook, copy sheet name).  This
        # allows the function to report the new sheet's name and index in the
        # sequence of sheets.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        copy_name = None
        for stored_name in self.spreadsheets.keys():
            if stored_name.lower() == sheet_name.lower():
                copy_name = stored_name
                break
        if not copy_name:
            raise KeyError(f"Sheet name {sheet_name} not found in workbook.")
        i = 1
        while True:
            copy_name = sheet_name + "_" + str(i)
            if copy_name.lower() not in self.spreadsheets.keys():
                break
            i += 1
            copy_name = copy_name[:-2]
        self.new_sheet(copy_name)
        for location, c in self.spreadsheets[stored_name].cells.items():
            self.set_cell_contents(copy_name, location, c.contents)
        return copy_name, len(self.spreadsheets) - 1
