from sheets import cell_error
import decimal


def str_to_error(str_error: str) -> cell_error.CellError:
    """
    Convert a string error to a cell error object.

    Args:
        str_error (str): string input

    Returns:
        cell_error.CellError: Cell Error object generated from the corresponding string form. 
    """
    match str_error.upper():
        case "#ERROR!":
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "input error")
        case "#CIRCREF!":
            return cell_error.CellError(cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
        case "#REF!":
            return cell_error.CellError(cell_error.CellErrorType.BAD_REFERENCE, "input error")
        case "#NAME?":
            return cell_error.CellError(cell_error.CellErrorType.BAD_NAME, "input error")
        case "#VALUE!":
            return cell_error.CellError(cell_error.CellErrorType.TYPE_ERROR, "input error")
        case "#DIV/0!":
            return cell_error.CellError(cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")
    return None


def is_number(string: str) -> bool:
    """
    Determine if a string can be represented as a number

    Args:
        string (str): input to test

    Returns:
        bool: if the string is numeric
    """
    string = string.strip()
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


def strip_evaluation(eval):
    """
    given evaluation from lark, return value
    """
    if type(eval) == str and is_number(eval):
        contents = strip_zeros(eval)
        return decimal.Decimal(contents)
    elif type(eval) == decimal.Decimal:
        stripped = strip_zeros(str(eval))
        return decimal.Decimal(stripped)
    else:
        return eval
