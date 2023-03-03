"""
Implementation of simple Function object and FunctionDirectory for the Workbook class. 
"""
from typing import List, Dict, Callable, Any
from decimal import Decimal
from sheets import cell_error, string_conversions, unitialized_value, version


class Function:
    """
    Representation of functions as a grouping of a string name and list of arguments.
    """

    def __init__(self, name: str, args: List[Any], lazy_eval: bool):
        self.name = name
        self.args = args
        self.lazy_eval = lazy_eval

    def __repr__(self):
        return f"Function: {self.name.strip()}{self.args}"


class FunctionDirectory:
    """
    Directory of callable functions for a workbook object. 
    """

    def __init__(self):
        # Function name -> callable function
        self.directory: Dict[str, Callable[[List[Any]], Any]] = {
            "AND": self.and_func,
            "OR": self.or_func,
            "NOT": self.not_func,
            "XOR": self.xor_func,
            "EXACT": self.exact_fn,
            "ISBLANK": self.is_blank,
            "ISERROR": self.is_error,
            "VERSION": self.version,
            "INDIRECT": self.indirect,
            "IF": self.if_func,
            "IFERROR": self.if_error,
            "CHOOSE": self.choose
        }
        self.lazy_functions = set(["IF", "IFERROR", "CHOOSE"])

    def get_function_keys(self):
        return self.directory.keys()

    def call_function(self, func: Function):
        try:
            return self.directory[func.name](func.args)
        except KeyError:
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "Invalid Function name")

    def and_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        found_false = False
        for arg in args:
            arg_eval = string_conversions.check_for_true_arg(arg)
            if isinstance(arg_eval, cell_error.CellError):
                return arg_eval
            if not arg_eval:
                found_false = True
        return not found_false

    def or_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        found_true = False
        for arg in args:
            arg_eval = string_conversions.check_for_true_arg(arg)
            if isinstance(arg_eval, cell_error.CellError):
                return arg_eval
            if arg_eval:
                found_true = True
        return found_true

    def not_func(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        arg = args[0]
        arg_eval = string_conversions.check_for_true_arg(arg)
        if isinstance(arg_eval, cell_error.CellError):
            return arg_eval
        return not arg_eval

    def xor_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        true_cnt = 0
        for arg in args:
            arg_eval = string_conversions.check_for_true_arg(arg)
            if isinstance(arg_eval, cell_error.CellError):
                return arg_eval
            if arg_eval:
                true_cnt += 1
        return true_cnt % 2 != 0

    def exact_fn(self, args: List):
        if len(args) != 2:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        error_found_0 = isinstance(args[0], cell_error.CellError)
        error_found_1 = isinstance(args[1], cell_error.CellError)
        if error_found_0:
            if error_found_1:
                return args[0] if args[0] < args[1] else args[1]
            return args[0]
        if error_found_1:
            return args[1]
        return str(args[0]) == str(args[1])

    def is_blank(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        value = args[0]
        if isinstance(value, cell_error.CellError):
            if value.get_type() == cell_error.CellErrorType.CIRCULAR_REFERENCE:
                return value
            return False
        if value == "" or value == Decimal(0):
            return False
        if isinstance(value, unitialized_value.UninitializedValue):
            return True
        if isinstance(value, bool) and value is False:
            return False
        return False

    def is_error(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        if isinstance(args[0], cell_error.CellError):
            if args[0].get_type() == cell_error.CellErrorType.CIRCULAR_REFERENCE:
                if args[0].circref_type:
                    args[0].circref_type = False
                    return args[0]
                if args[0].circref_type is False or args[0].circref_type is None:
                    return True
            return True
        return False

    def version(self, args: List):
        if len(args) != 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        return version.version # pylint: disable=no-member

    def indirect(self, args: List):
        return args[0]

    def if_func(self, args: List):
        return args[0]

    def if_error(self, args: List):
        return args[0]

    def choose(self, args: List):
        return args[0]
