"""Workbook API. Contains spreadsheet functions accessible to public users."""
from __future__ import annotations
from typing import Tuple, List, Optional, Any, TextIO, Callable, Iterable, Dict, Set
import decimal
import copy
import json
import re
from contextlib import contextmanager, suppress
from sheets import cell, topo_sort, cell_error, lark_module, sheet, \
    string_conversions, unitialized_value


ALLOWED_PUNC = set([".", "?", "!", ",", ":", ";", "@", "#",
                    "$", "%", "^", "&", "*", "(", ")", "-", "_"])


class Workbook:
    """
    A workbook containing zero or more named spreadsheets.

    Any and all operations on a workbook that may affect calculated cell
    values should cause the workbook's contents to be updated properly.
    """

    def __init__(self):
        # lower case name -> sheet object
        self.spreadsheets: Dict[str, sheet.Sheet] = {}
        # Cell: [neighbor Cells]; neighbors are cells that depend on Cell
        self.adjacency_list: Dict[cell.Cell, List[cell.Cell]] = {}
        # notify functions = set of user-inputted notify functions
        self.notify_functions: List[Callable[[
            Workbook, Iterable[Tuple[str, str]]], None]] = []
        # Whether we should call __generate_notifications() in set_cell_contents()
        # Pertinent in move_cells() and copy_cells() where cells are updated multiple times
        self.__call_notify = True

    @contextmanager
    def __disable_notify_calls(self):
        """
        Disable automatic calls to notify functions in set_cell_contents
        """
        self.__call_notify = False
        yield
        self.__call_notify = True

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
        if sheet_name[-1] == " ":
            raise ValueError("Sheet name ends with white space")
        if sheet_name.lower() in self.spreadsheets:
            raise ValueError("Duplicate spreadsheet name")

    def __update_extent(self, spreadsheet: sheet.Sheet, location: str, deleting_cell: bool):
        """
        Update the extent of a sheet if we are deleting a cell
        """
        if deleting_cell:
            sheet_col, sheet_row = spreadsheet.extent_col, spreadsheet.extent_row
            col, row = string_conversions.str_to_tuple(location)
            max_col, max_row = 0, 0
            if col == sheet_col or row == sheet_row:
                for c in spreadsheet.cells:
                    if spreadsheet.cells[c].cell_type != cell.CellType.EMPTY:
                        c_col, c_row = string_conversions.str_to_tuple(
                            spreadsheet.cells[c].location)
                        max_col = max(max_col, c_col)
                        max_row = max(max_row, c_row)
                spreadsheet.extent_col = max_col
                spreadsheet.extent_row = max_row
        else:
            curr_col, curr_row = string_conversions.str_to_tuple(location)
            spreadsheet.extent_col = max(curr_col, spreadsheet.extent_col)
            spreadsheet.extent_row = max(curr_row, spreadsheet.extent_row)

    def __generate_notifications(self, cell_list: Iterable[cell.Cell]):
        """Given a list of cells, create a corresponding list of tuples containing
        each cell's sheet name and location. Then, call each registered notify
        function on the created list.

        Args:
            cell_list (Iterable[Cell]): Collection of cell obejcts to call ntoify functions on.
        """
        if cell_list == [] or len(cell_list) == 0:
            return
        changed_cells = [(c.sheet.name, c.location) for c in cell_list]
        for notify_function in self.notify_functions:
            with suppress(Exception):
                notify_function(copy.deepcopy(self), changed_cells)

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
        # determine the new type of our cell and set its value accordingly
        if not cell_contents or len(cell_contents) == 0:
            val = None
            cell_type = cell.CellType.EMPTY
        elif string_conversions.str_to_error(cell_contents):  # type error
            val = string_conversions.str_to_error(cell_contents)
            cell_type = cell.CellType.ERROR
        elif cell_contents[0] == "'":
            val = cell_contents[1:]
            cell_type = cell.CellType.STRING
        elif cell_contents[0] == "=":
            evaluator, val = lark_module.evaluate_expr(
                self, calling_cell, calling_cell.sheet.name, cell_contents)
            cell_type = cell.CellType.FORMULA
            if evaluator:
                relies_on = evaluator.calling_cell_relies_on
        elif string_conversions.is_number(cell_contents):
            stripped = string_conversions.strip_zeros(cell_contents)
            val = decimal.Decimal(stripped)
            cell_type = cell.CellType.LITERAL_NUM
        else:
            val = cell_contents
            cell_type = cell.CellType.LITERAL_STRING
        # determine if updating the value actually updates it or changes its type
        type_change = False
        if calling_cell.cell_type == cell.CellType.EMPTY:
            if val != unitialized_value.UninitializedValue:
                type_change = True
        elif calling_cell.cell_type != cell_type:
            type_change = True
        val_update = bool(val != calling_cell.value or type_change)
        calling_cell.set_fields(value=val, cell_type=cell_type)
        return relies_on, val_update

    def __get_cells_containing_sheetname(self, sheetname: str) -> list[cell.Cell]:
        # match any cell that has contents sheetname! or 'sheetname'!
        cells = []
        regex = f'.*{sheetname.lower()}!.*'
        regex2 = f".*'{sheetname.lower()}'!.*"
        for name, spreadsheet in self.spreadsheets.items():
            for loc, c in spreadsheet.cells.items():
                if c.contents and \
                    (re.match(regex, c.contents.lower()) or
                        re.match(regex2, c.contents.lower())):
                    try:
                        # If a cell is an error type, check if it is a parse error and don't
                        # add to our list of cells. Otherwise, if it is not a cell error type,
                        # get_type() will throw an attribute error.
                        if self.get_cell_value(name, loc).get_type() == \
                                cell_error.CellErrorType.PARSE_ERROR:
                            continue
                    except AttributeError:
                        pass
                    cells.append(c)
        return cells

    def __get_cell_contents_after_rename(self, c: cell.Cell, sheet_name: str,
                                         new_sheet_name: str) -> str:
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

    def __get_selection_corners(self, start_location: str,
                                end_location: str) -> Tuple[int, int, int, int]:
        """
        Given to corners of our selection, return the locations of the top left and bottom
        right of our selection regardless of the input location

        Args:
            start_location (str): Start of our selection
            end_location (str): End of our selection

        Returns:
            Tuple[int, int, int, int]: integers representing the top left column, top left row, 
            bottom right column, and bottom right row
        """
        start_col, start_row = string_conversions.str_to_tuple(start_location)
        end_col, end_row = string_conversions.str_to_tuple(end_location)
        top_left_col = min(start_col, end_col)
        top_left_row = min(start_row, end_row)
        bottom_right_col = max(start_col, end_col)
        bottom_right_row = max(start_row, end_row)
        return top_left_col, top_left_row, bottom_right_col, bottom_right_row

    def __get_overlap_map(self, spreadsheet: sheet.Sheet,
                          original_corners: Tuple[int, int, int, int],
                          destination_corners: Tuple[int, int, int, int]
                          ) -> Dict[str, Tuple[str, cell.CellType]]:
        """Get mapping of cells in overlapping region to their original contents.
        Used in move_cells and copy_cells.

        Args:
            original_corners (Tuple[int, int, int, int]): 
            (top left col, top left row, bottom right col, bottom right row)
            destination_corners (Tuple[int, int, int, int]):
            (top left col, top left row, bottom right col, bottom right row)

        Returns:
            Dict[str, Tuple[str, cell.CellType]]: Mapping of location to contents and cell type
        """
        mapping = {}
        for i in range(destination_corners[0], destination_corners[2] + 1):
            for j in range(destination_corners[1], destination_corners[3] + 1):
                if original_corners[0] <= i <= original_corners[2] and \
                        original_corners[1] <= j <= original_corners[3]:
                    loc = "".join([string_conversions.num_to_col(i), str(j)])
                    contents = self.get_cell_contents(spreadsheet.name, loc)
                    cell_type = spreadsheet.cells[loc].cell_type
                    mapping[loc] = (contents, cell_type)
        return mapping

    def __copy_cell_block(self, spreadsheet: sheet.Sheet, start_location: str,
                          end_location: str, to_location: str, to_sheet: str,
                          deleting: bool) -> Set[cell.Cell]:
        """
        Copy a block of cells from one location to another

        Args:
            spreadsheet (Sheet): Sheet object our copied cells are located in
            sheet_name (str): Name of the sheet our starting cells are in
            start_location (str): One corner of our selection block
            end_location (str): Opposite corner of our selection block
            to_location (str): Top-left corner of the region we intend to place our cells in
            to_sheet (str): Sheet to place our cell block in
            deleting (bool): Whether we are deleting our original block of cells after moving them

        Returns:
            Set[Cell]: iterable set of all cell objects whose values are changed by the copying
            operation. 
        """
        affected_cells = set()
        start_location = start_location.upper()
        end_location = end_location.upper()
        to_location = to_location.upper()
        end_top_left_col, end_top_left_row = string_conversions.str_to_tuple(
            to_location)
        top_left_col, top_left_row, bottom_right_col, bottom_right_row = \
            self.__get_selection_corners(start_location, end_location)
        # get the required change in column and row
        to_col, to_row = string_conversions.str_to_tuple(to_location)
        delta_col = to_col - top_left_col
        delta_row = to_row - top_left_row
        # obtain mapping of overlapped cells to their original content
        end_bottom_right_col = bottom_right_col + delta_col
        end_bottom_right_row = bottom_right_row + delta_row
        original_corners = (top_left_col, top_left_row,
                            bottom_right_col, bottom_right_row)
        destination_corners = (
            end_top_left_col, end_top_left_row, end_bottom_right_col, end_bottom_right_row)
        # if our top left corner is equal to to_location, then we aren't actually moving any
        # cells and we can just return an empty set.
        if "".join(
            [string_conversions.num_to_col(top_left_col), str(top_left_row)]) == to_location:
            return affected_cells
        overlap_map = self.__get_overlap_map(
            spreadsheet, original_corners, destination_corners)
        # move each cell in our selection zone
        for i in range(top_left_col, bottom_right_col + 1):
            start_cell_col = string_conversions.num_to_col(i)
            end_cell_col = string_conversions.num_to_col(i + delta_col)
            for j in range(top_left_row, bottom_right_row + 1):
                start_cell_loc = start_cell_col + str(j)
                end_cell_loc = end_cell_col + str(j + delta_row)
                # If end cell location in overlap region, get its original contents
                if start_cell_loc in overlap_map:
                    contents = overlap_map[start_cell_loc][0]
                else:
                    contents = self.get_cell_contents(
                        spreadsheet.name, start_cell_loc)
                # If cell is formula, we need to update its relative location for cell references
                # It's possible that a cell in the overlap region is overwritten from formula to
                # string so we need to check its original cell type
                overwritten_formula = start_cell_loc in overlap_map and overlap_map[
                    start_cell_loc][1] == cell.CellType.FORMULA
                originally_formula = start_cell_loc in spreadsheet.cells and \
                    spreadsheet.cells[
                        start_cell_loc].cell_type == cell.CellType.FORMULA
                if overwritten_formula or originally_formula:
                    # If a cell is an error type, check if it is a parse error and don't
                    # update its contents. Otherwise, if it is not a cell error type,
                    # get_type() will throw an attribute error.
                    try:
                        if self.get_cell_value(spreadsheet.name, start_cell_loc).get_type() == \
                                cell_error.CellErrorType.PARSE_ERROR:
                            self.set_cell_contents(
                                to_sheet, end_cell_loc, contents)
                            affected_cells.add(
                                self.spreadsheets[to_sheet.lower()].cells[end_cell_loc])
                            continue
                    except AttributeError:
                        pass
                    # Since cell type is not error, we don't worry about invalid cell refs
                    # Locations can be specified as A1, sheet1!A1, or 'sheet1'!A1
                    # sheetname: \'[^']*\'! OR [A-Za-z_][A-Za-z0-9_]*!
                    sheetname_pattern = r"\'[^']*\'!|[A-Za-z_][A-Za-z0-9_]*!"
                    cell_pattern = r"\$?[A-Za-z]+\$?[1-9][0-9]*"
                    # ?: Specifies we don't want to keep the matched sheetname
                    pattern = f"(?:{sheetname_pattern})?({cell_pattern})"
                    locations = re.findall(pattern, contents)
                    for loc in locations:
                        loc = loc.upper()
                        # If $ precedes col or row, do not update relative location
                        # In the regex, () defines the two groups
                        match = re.match(
                            r"(\$?[A-Za-z]+)(\$?[1-9][0-9]*)", loc)
                        col = match.group(1)
                        row = match.group(2)
                        new_loc = ""
                        if col[0] != "$":
                            col = string_conversions.col_to_num(
                                col) + delta_col
                            new_loc += string_conversions.num_to_col(col)
                        else:
                            new_loc += col
                        if row[0] != "$":
                            row = delta_row + int(row)
                            new_loc += str(row)
                        else:
                            new_loc += row
                        if not spreadsheet.check_valid_location(new_loc):
                            new_loc = "#REF!"
                        # Note that re.escape ensures we can properly search for $ in loc
                        # Normally, you'd have to escape $A$1 like \$A\$1
                        contents = re.sub(
                            re.escape(loc), new_loc, contents, flags=re.IGNORECASE)
                    self.set_cell_contents(to_sheet, end_cell_loc, contents)
                    affected_cells.add(
                        self.spreadsheets[to_sheet.lower()].cells[end_cell_loc])
                # Cells that aren't formulas can copy the original location's contents
                else:
                    self.set_cell_contents(
                        to_sheet, end_cell_loc, contents)
                    affected_cells.add(
                        self.spreadsheets[to_sheet.lower()].cells[end_cell_loc])
                if deleting:
                    # Delete cells that don't overlap with the new location
                    if start_cell_loc not in overlap_map:
                        affected_cells.add(spreadsheet.cells[start_cell_loc])
                        self.set_cell_contents(
                            spreadsheet.name, start_cell_loc, None)
        return affected_cells

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
        return [self.spreadsheets[spreadsheet].name for spreadsheet in self.spreadsheets]

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
        if sheet_name:
            self.__check_valid_sheet_name(sheet_name)
        else:  # handle null input
            i = 1
            while True:
                sheet_name = "Sheet" + str(i)
                if sheet_name.lower() not in self.spreadsheets:
                    break
                i += 1
        self.spreadsheets[sheet_name.lower()] = sheet.Sheet(sheet_name)
        changed_cells = [c for c in list(self.adjacency_list.keys())
                         if self.__set_cell_value_and_type(c)[1]]
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
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        del self.spreadsheets[sheet_name.lower()]
        changed_cells = []
        for loc in spreadsheet.cells:
            c = spreadsheet.cells[loc]
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
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        return ((spreadsheet.extent_col, spreadsheet.extent_row))

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
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        if not spreadsheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if contents:
            contents = contents.strip()
        # if cell already exists (modify contents)
        if location in spreadsheet.cells:
            existing_cell = spreadsheet.cells[location]
            existing_cell.contents = contents
            relies_on, val_updated = self.__set_cell_value_and_type(
                existing_cell)
            # Everything the existing cell relies on
            for c, neighbors in self.adjacency_list.items():
                if existing_cell in neighbors and c not in relies_on:
                    neighbors.remove(existing_cell)
            if existing_cell.cell_type == cell.CellType.EMPTY:
                # existing cell is now empty so it does not depend on other cells
                # -> remove it as a neighbor of other cells
                for _, neighbors in self.adjacency_list.items():
                    if existing_cell in neighbors:
                        neighbors.remove(existing_cell)
                # if existing cell doesn't have neighbors, no cell relies on it
                # -> delete cell from spreadsheet
                if not self.adjacency_list[existing_cell]:
                    del spreadsheet.cells[location]
                    del self.adjacency_list[existing_cell]
                    self.__update_extent(spreadsheet, location, True)
                    if self.__call_notify:
                        self.__generate_notifications([existing_cell])
                    return
            # updating cells that depend on existing cell only if the value is updated
            if val_updated:
                circular, cell_dependents = topo_sort(
                    existing_cell, self.adjacency_list)
                if not circular:
                    for dependent in cell_dependents[1:]:
                        self.__set_cell_value_and_type(dependent)
                else:  # everything in the cycle should have an error value
                    for dependent in cell_dependents:
                        dependent.set_fields(value=cell_error.CellError(
                            cell_error.CellErrorType.CIRCULAR_REFERENCE, "circular reference"))
                self.__update_extent(spreadsheet, location, False)
                # include the existing cell iff its value is updated
                if self.__call_notify:
                    self.__generate_notifications(cell_dependents)
        else:  # if cell does not exist (create contents)
            new_cell = cell.Cell(spreadsheet, location, contents, None, None)
            self.__set_cell_value_and_type(new_cell)
            if new_cell.cell_type != cell.CellType.EMPTY:
                self.adjacency_list[new_cell] = []
                spreadsheet.cells[location] = new_cell
            self.__update_extent(spreadsheet, location, False)
            if self.__call_notify:
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
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        if not spreadsheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if location in spreadsheet.cells:
            return spreadsheet.cells[location].contents
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
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        if not spreadsheet.check_valid_location(location):
            raise ValueError(f"Cell location {location} is invalid")
        if location in spreadsheet.cells:
            if isinstance(spreadsheet.cells[location].value, unitialized_value.UninitializedValue):
                return decimal.Decimal(0)
            return spreadsheet.cells[location].value
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
        if not isinstance(sheets, list):
            raise TypeError("\"sheets\" value should be of type list")

        for spreadsheet in sheets:
            if not isinstance(spreadsheet, dict):
                raise TypeError("Sheet object is not of type dictionary.")
            if len(spreadsheet.keys()) != 2:
                raise KeyError(
                    "Sheet dictionary does not have exactly two keys.")

            if "name" not in spreadsheet.keys() or "cell-contents" not in spreadsheet.keys():
                raise KeyError(
                    "Keys of sheet dictionray must be named \"name\" and \"cell-contents\"")
            sheet_name = spreadsheet["name"]
            if not isinstance(sheet_name, str):
                raise TypeError("Sheet name is not of type string.")
            wb.new_sheet(sheet_name)

            cell_contents = spreadsheet["cell-contents"]
            if not isinstance(cell_contents, dict):
                raise TypeError(
                    f"Cell Contents of sheet \"{sheet_name}\" is not of type dict.")
            for location, contents in cell_contents.items():
                if not isinstance(location, str) or not isinstance(contents, str):
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
        for _, spreadsheet in self.spreadsheets.items():
            name = spreadsheet.name
            cur_sheet = {'name': name, 'cell-contents': {}}
            for location, c in spreadsheet.cells.items():
                cur_sheet['cell-contents'][location] = c.contents
            data["sheets"].append(cur_sheet)
        try:
            json.dump(data, fp, indent=4)
        except IOError as e:
            print(f"{e}, IO write error in save_workbook().")

    def notify_cells_changed(self, notify_function:
                             Callable[[Workbook, Iterable[Tuple[str, str]]],
                                      None]) -> None:
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
        stored_name = None
        for stored_name in self.spreadsheets:
            if stored_name.lower() == sheet_name.lower():
                copy_name = stored_name
                break
        if not copy_name:
            raise KeyError(f"Sheet name {sheet_name} not found in workbook.")
        i = 1
        while True:
            copy_name = sheet_name + "_" + str(i)
            if copy_name.lower() not in self.spreadsheets:
                break
            i += 1
            copy_name = copy_name[:-2]
        self.new_sheet(copy_name)
        for location, c in self.spreadsheets[stored_name].cells.items():
            self.set_cell_contents(copy_name, location, c.contents)
        return copy_name, len(self.spreadsheets) - 1

    def move_cells(self, sheet_name: str, start_location: str,
                   end_location: str, to_location: str, to_sheet: Optional[str] = None) -> None:
        # Move cells from one location to another, possibly moving them to
        # another sheet.  All formulas in the area being moved will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) will
        # become empty due to the move operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be moved.  The to_location specifies the
        # top-left corner of the target area to move the cells to.
        #
        # Both corners are included in the area being moved; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to move, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to move.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being moved to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being moved contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError(f"{sheet_name} is invalid")
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        for location in [start_location, end_location, to_location]:
            if not spreadsheet.check_valid_location(location):
                raise ValueError(f"Cell location {location} is invalid")
        if not to_sheet:
            to_sheet = sheet_name
        with self.__disable_notify_calls():
            affected_cells = self.__copy_cell_block(spreadsheet, start_location,
                                                    end_location, to_location, to_sheet, True)
        self.__generate_notifications(affected_cells)

    def copy_cells(self, sheet_name: str, start_location: str,
                   end_location: str, to_location: str, to_sheet: Optional[str] = None) -> None:
        # Copy cells from one location to another, possibly copying them to
        # another sheet.  All formulas in the area being copied will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) are
        # left unchanged by the copy operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be copied.  The to_location specifies the
        # top-left corner of the target area to copy the cells to.
        #
        # Both corners are included in the area being copied; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to copy, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to copy.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being copied to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being copied contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        if sheet_name.lower() not in self.spreadsheets:
            raise KeyError(f"{sheet_name} is invalid")
        spreadsheet = self.spreadsheets[sheet_name.lower()]
        for location in [start_location, end_location, to_location]:
            if not spreadsheet.check_valid_location(location):
                raise ValueError(f"Cell location {location} is invalid")
        if not to_sheet:
            to_sheet = sheet_name
        with self.__disable_notify_calls():
            affected_cells = self.__copy_cell_block(spreadsheet, start_location,
                                                    end_location, to_location, to_sheet, False)
        self.__generate_notifications(affected_cells)
