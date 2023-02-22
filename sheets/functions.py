from typing import List, Dict, Callable
from decimal import Decimal
from sheets import cell_error, string_conversions


class FunctionDirectory:## func_name : callable function
    def __init__(self):
        self.directory: Dict[str, Callable] = {
            "AND": and_func
        }

    def call_function(self, func_name: str, args: List):
        return self.directory[func_name](args)


def and_func(args: List):
    for arg in args:
        if isinstance(arg, bool) and not arg:
            return False
        elif isinstance(arg, Decimal):
            if arg == Decimal(0):
                return False
        elif string_conversions.is_number(arg):
            if Decimal(arg) == Decimal(0):
                return False
        elif isinstance(arg, str):
            if arg.lower() == "false":
                return False
            if arg.lower() != "true":
                return cell_error.CellError(
                    cell_error.CellErrorType.TYPE_ERROR, "Invalid string argument")
    return True 
