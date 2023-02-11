"""
Utility functions and helpers for test files.
"""

import sys
import io
from typing import Tuple, TextIO


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
