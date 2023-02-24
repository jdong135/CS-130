"""
Module containing functions involving converting strings between ints,
errors, and numbers. Also contains method to strip zeros from string num. 
"""
import decimal
from typing import Tuple
from functools import cache
from sheets import cell_error


ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
COMPARISON_OPERATORS = ["=", "==", "<>", "!=", ">", "<", ">=", "<="]


@cache
def str_to_tuple(location: str) -> Tuple[int, int]:
    """
    Take in a string location ranging from A1 to ZZZZ9999 and return the 
    integer coordinates of the location.

    Args:
        location (str): string location on the sheet.

    Returns:
        Tuple[int, int]: numeric coordinates on the sheet equivalent to 
        input location. 
    """
    # If location has absolute references, we can't parse through them
    location = location.replace("$", "")
    chars = list(location)
    rows = 0
    cols = 0
    rowCnt = 0
    colCnt = 0
    for i in range(len(chars) - 1, -1, -1):
        if chars[i].isnumeric():
            rows += (10 ** rowCnt) * int(chars[i])
            rowCnt += 1
        else:
            cols += (26 ** colCnt) * (ord(chars[i]) - ord('A') + 1)
            colCnt += 1
    return (cols, rows)


def col_to_num(location: str) -> Tuple[int, int]:
    """
    Take in a string column ranging from A to ZZZZ and return the 
    integer equivalent of the location.

    Args:
        location (str): string location on the sheet.

    Returns:
        (int): numeric coordinates on the sheet equivalent to 
        input location. 
    """
    chars = list(location)
    cols = 0
    colCnt = 0
    for i in range(len(chars) - 1, -1, -1):
        cols += (26 ** colCnt) * (ord(chars[i]) - ord('A') + 1)
        colCnt += 1
    return cols


def num_to_col(col: int) -> str:
    """
    Convert an integer representation of a row to its corresponding string

    Args:
        col (int): Location to convert to string

    Returns:
        str: String representation of input location
    """
    res = []
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        res.append(ALPHABET[remainder])
    return "".join(reversed(res))


def str_to_error(str_error: str) -> cell_error.CellError:
    """
    Convert a string error to a cell error object.

    Args:
        str_error (str): string input

    Returns:
        cell_error.CellError: Cell Error object generated from the corresponding string form. 
    """
    err = None
    match str_error.upper():
        case "#ERROR!":
            err = cell_error.CellError(
                cell_error.CellErrorType.PARSE_ERROR, "input error")
        case "#CIRCREF!":
            err = cell_error.CellError(
                cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
        case "#REF!":
            err = cell_error.CellError(
                cell_error.CellErrorType.BAD_REFERENCE, "input error")
        case "#NAME?":
            err = cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "input error")
        case "#VALUE!":
            err = cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "input error")
        case "#DIV/0!":
            err = cell_error.CellError(
                cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")
    return err


def is_number(string: str) -> bool:
    """
    Determine if a string can be represented as a number

    Args:
        string (str): input to test

    Returns:
        bool: if the string is numeric
    """
    string = string.strip()
    if string[0] == "-":
        return string[1:].isnumeric() or (
            string[1:].replace('.', '', 1).isdigit() and string[1:].count('.') < 2)
    return string.isnumeric() or (string.replace('.', '', 1).isdigit() and string.count('.') < 2)


def strip_zeros(string: str) -> str:
    """
    Strip all trailing zeros from a string representation of a number

    Args:
        string (str): string to remove zeros from 

    Returns:
        str: representation of number without trailing zeros
    """
    return string.rstrip('0').rstrip('.') if '.' in string else string


def strip_evaluation(evaluation):
    """
    given evaluation from lark, return value
    """
    if isinstance(evaluation, str) and is_number(evaluation):
        contents = strip_zeros(evaluation)
        return decimal.Decimal(contents)
    if isinstance(evaluation, decimal.Decimal):
        stripped = strip_zeros(str(evaluation))
        return decimal.Decimal(stripped)
    return evaluation

def is_bool_expr(expr: str):
    if expr.lower() == "false" or expr.lower() == "true":
        return True
    else:
        return False

def is_true_expr(expr: str):
    if expr.lower() == "true":
        return True
    else:
        return False
