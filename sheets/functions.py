from typing import List, Dict, Callable, Any, Tuple, Union
from decimal import Decimal
from sheets import cell_error, string_conversions, unitialized_value, version
import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class Function:
    """
    Representation of functions as a grouping of a string name and list of arguments.
    """
    def __init__(self, name: str, args: List[Any]):
        self.name = name
        self.args = args

    def __repr__(self):
        return f"Function: {self.name.strip()}{self.args}"

    @staticmethod
    def parse_function_by_index(func_call: str) -> Tuple[str, List[Any], int]:
        """
        Given a function of the form "FUNC(arg1, ..., argn)", parse the function
        into a corresponding function name and list of arguments that can be
        fed into the functions.Function() constructor. Note that this will
        recursively create function objets out of arguments that take the form of
        a parsable function. 

        Args:
            func_call (str): String function to parse

        Returns:
            Tuple[str, List[Any], int]: Function name, list of corresponding arguments,
            and amount of charactes parsed on the given recursive layer.
        """
        accumulator = ""
        head = ""
        i = 0
        writing_head = True
        while True:
            curr = func_call[i]
            if curr == "(":
                if writing_head:
                    head = accumulator
                    accumulator = ""
                    args = []
                    writing_head = False
                else:
                    sub_head = accumulator
                    accumulator = ""
                    _, sub_args, d = Function.parse_function_by_index(func_call[i:])
                    i += d
                    args.append(Function(sub_head.strip().upper(), sub_args))
            elif curr == ",":
                accumulator = accumulator.strip()
                if len(accumulator) > 0:
                    args.append(accumulator)
                    accumulator = ""
            elif curr == ")":
                accumulator = accumulator.strip()
                if len(accumulator) > 0:
                    args.append(accumulator)
                    accumulator = ""
                break
            else:
                accumulator += curr
            i += 1
        return head, args, i


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
            "IFERROR": self.if_error
        }

    def call_function(self, func_name: str, args: List):
        try:
            return self.directory[func_name](args)
        except KeyError:
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "Invalid Function name")

    def and_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        found_false = False
        for arg in args:
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
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
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
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
        if isinstance(arg, Function):
            arg = self.call_function(arg.name, arg.args)
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
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
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
        return str(args[0]) == str(args[1])

    def is_blank(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        value = args[0]
        if value == "" or value == Decimal(0):
            return False
        if isinstance(value, unitialized_value.UninitializedValue):
            return True
        if isinstance(value, bool) and value == False:
            return False
        return False

    def is_error(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        if isinstance(args[0], cell_error.CellError):
            return True
        return False

    def version(self, args: List):
        if len(args) != 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        return version.version

    def indirect(self, args: List):
        return args[0]

    def if_func(self, args: List):
        logger.info(args)
        return args[1] if args[0] else args[2]

    def if_error(self, args: List):
        logger.info(args)
        if isinstance(args[0], cell_error.CellError):
            return args[1]
        return args[0]
