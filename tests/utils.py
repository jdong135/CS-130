"""
Utility functions and helpers for test files.
"""

import sys
import io
import re
import decimal
from typing import Tuple, TextIO, List, Any
from context import sheets


def store_stdout() -> Tuple[io.StringIO, TextIO]:
    """
    Generate a 'dummy' copy of stdout that we can stores console outputs
    and create a variable containing the original sys.stdout.

    Returns:
        Tuple[io.StringIO, TextIO]: A tuple containing the StringIO in which
        all proceeding console outputs will be redirected into and a TextIO 
        variable that contains the file analgous to Python's interpreter's 
        standard output stream. 
    """
    sys_out = sys.stdout
    new_stdo = io.StringIO()
    sys.stdout = new_stdo
    return new_stdo, sys_out


def restore_stdout(new_stdo: io.StringIO, sys_out: TextIO) -> str:
    """
    Restore the original stdout and return the contents stored in our temporary
    variable

    Args:
        new_stdo (io.StringIO): Temporary variable storing all contents
        directed into stdout. 
        sys_out (TextIO): The original file analgous to Python's standard
        output stream. 

    Returns:
        str: Contents stored in the `new_stdo` variable
    """
    output = new_stdo.getvalue()
    sys.stdout = sys_out
    return output


def sort_notify_list(notify_list: str) -> List[str]:
    """
    Sort a string list generated by a notify function that prints its input of changed cells
    to the console. Since sets are ordered differently on each iteration, this is necessary
    to ensure our ntoify function calls the same list of cells every time when moving and
    copying cells. 

    Args:
        notify_list (str): Console output of list of cells from a notify funciton

    Returns:
        List[str]: Sorted substrings of notify_list 
    """
    matches = re.findall(r"\((.*?)\)", notify_list)
    matches.sort()
    return matches


def on_cells_changed(workbook, cells_changed):
    """Standard notification function."""
    _ = workbook
    print(cells_changed)

def block_equal(wb, unittest, vals: List[Any], sheet_name: str=None):
    """
    Assert that a block of values in the A-B-C columns are equal to a provided list
    of values, which are ordered as (A1, B1, C1, A2, B2, ...). If a value is a CellError,
    compare with its type.

    Args:
        wb (Workbook): Instance of workbook class
        unittest (WorkbookSortCells): Instance of unittesting class
        vals (List[Any]): list of values to check
    """
    cols_by_mod = {0: 'A', 1: 'B', 2: 'C'}
    if not sheet_name:
        sheet_name = "sheet1"
    for i, _ in enumerate(vals):
        loc = str(cols_by_mod[i % 3]) + str((i // 3) + 1)
        value = wb.get_cell_value(sheet_name, loc)
        if isinstance(value, sheets.cell_error.CellError):
            value = value.get_type()
        if isinstance(vals[i], (int, float)):
            vals[i] = decimal.Decimal(vals[i])
        unittest.assertEqual(value, vals[i])
