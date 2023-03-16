import copy
import re
from decimal import Decimal
from typing import List
from sheets import cell, cell_error, sheet, \
    string_conversions, unitialized_value, row

ALLOWED_PUNC = set([".", "?", "!", ",", ":", ";", "@", "#",
                    "$", "%", "^", "&", "*", "(", ")", "-", "_"])


def check_valid_sheet_name(wb, sheet_name: str) -> None:
    """
    Check if a sheet name is valid.

    Args:
        wb (Workbook): workbook instance
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
    if sheet_name.lower() in wb.spreadsheets:
        raise ValueError("Duplicate spreadsheet name")
    

def update_extent(spreadsheet: sheet.Sheet, location: str, deleting_cell: bool):
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


def create_row_list(top_left_col, top_left_row, bottom_right_col, bottom_right_row, 
                    spreadsheet, sort_cols) -> List[row.Row]:
    row_list = []
    for i in range(top_left_row, bottom_right_row + 1):
        cur_row = []  # list of Cell objects
        for j in range(top_left_col, bottom_right_col + 1):
            start_cell_col = string_conversions.num_to_col(j)
            cell_loc = start_cell_col + str(i)
            if cell_loc in spreadsheet.cells:
                deep_copy = copy.deepcopy(spreadsheet.cells[cell_loc])
                cur_row.append(deep_copy)
            else:
                cur_row.append(unitialized_value.EmptySortCell())
        row_list.append(row.Row(sort_cols, i, cur_row))
    return row_list


def update_all_block_contents(wb, sheet_name, row_list, top_left_col, top_left_row) -> None:
    for i, row_obj in enumerate(row_list):
        for j, c in enumerate(row_obj.cell_list):
            new_row = top_left_row + i
            new_col = string_conversions.num_to_col(top_left_col + j)
            new_location = new_col + str(new_row)
            if not isinstance(
                c, unitialized_value.EmptySortCell) and c.cell_type == cell.CellType.FORMULA:
                _, old_row = string_conversions.str_to_tuple(
                    c.location)
                delta_row = new_row - old_row
                sheetname_pattern = r"\'[^']*\'!|[A-Za-z_][A-Za-z0-9_]*!"
                cell_pattern = r'(?<!")\$?[A-Za-z]+\$?[1-9][0-9]*(?!")'
                pattern = f"(?:{sheetname_pattern})?({cell_pattern})"
                locations = re.findall(pattern, c.contents)
                for loc in locations:
                    loc = loc.upper()
                    # If $ precedes col or row, do not update relative location
                    # In the regex, () defines the two groups
                    match = re.match(
                        r"(\$?[A-Za-z]+)(\$?[1-9][0-9]*)", loc)
                    col_match = match.group(1)
                    row_match = match.group(2)
                    new_loc = col_match
                    if row_match[0] != "$":
                        # row_match = top_left_row + i
                        row_match = delta_row + int(row_match)
                        new_loc += str(row_match)
                    else:
                        new_loc += row_match
                    if not string_conversions.check_valid_location(new_loc):
                        new_loc = "#REF!"
                    # Note that re.escape ensures we can properly search for $ in loc
                    # Normally, you'd have to escape $A$1 like \$A\$1
                    c.contents = re.sub(
                        re.escape(loc), new_loc, c.contents, flags=re.IGNORECASE)
            if not isinstance(c, unitialized_value.EmptySortCell):
                wb.set_cell_contents(
                    sheet_name.lower(), new_location, c.contents)
                

def compare(obj1, obj2):
    # boolean > str > decimal > error > uninitialized (blank)
    type_map = {
        unitialized_value.EmptySortCell: 1,
        unitialized_value.UninitializedValue : 1,
        type(None) : 1,
        cell_error.CellError: 2,
        Decimal: 3,
        str: 4,
        bool: 5
    }
    for i, sort_col in enumerate(obj1.sort_cols):
        index = abs(sort_col) - 1
        item1, item2 = obj1.cell_list[index], obj2.cell_list[index]
        if isinstance(item1, cell.Cell):
            item1 = item1.value
        if isinstance(item2, cell.Cell):
            item2 = item2.value
        val1 = type_map[type(item1)]
        val2 = type_map[type(item2)]
        reverse = obj1.rev_list[i]
        return_val = None
        if val1 == val2:
            if val1 == 2:
                if item1._error_type.value < item2._error_type.value:
                    return_val = -1
                elif item1._error_type.value > item2._error_type.value:
                    return_val = 1
            elif val1 != 1:
                if item1 < item2:
                    return_val = -1
                elif item1 > item2:
                    return_val = 1
        else:
            if val1 < val2:
                return_val = -1
            elif val1 > val2:
                return_val = 1

        # If return value isn't null, return -1 or 1
        if return_val and reverse:
            return -1 * return_val
        elif return_val:
            return return_val
    # All checks for inequality fail, two lists must be equal
    return 0
