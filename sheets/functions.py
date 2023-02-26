from typing import List, Dict, Callable, Any
from decimal import Decimal
from sheets import cell_error, string_conversions, unitialized_value, version
import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class Function:
    def __init__(self, name: str, args: List[Any]):
        self.name = name
        self.args = args

    def __repr__(self):
        return f"Function: {self.name.strip()}{self.args}"


def parse_function_by_index(func_call: str):
    """
    deal with later

    Args:
        func_call (str): _description_

    Returns:
        _type_: _description_
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
                _, sub_args, d = parse_function_by_index(func_call[i:])
                i += d
                args.append(Function(sub_head.strip(), sub_args))
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
        }

    def call_function(self, func_name: str, args: List):
        try:
            return self.directory[func_name.upper()](args)
        except KeyError:
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "Invalid Function name")

    def and_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        for arg in args:
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
            if isinstance(arg, unitialized_value.UninitializedValue):
                return False
            if isinstance(arg, bool) and not arg:
                return False
            if isinstance(arg, cell_error.CellError):
                if arg.get_type() == cell_error.CellErrorType.BAD_REFERENCE:
                    return cell_error.CellError(
                        cell_error.CellErrorType.BAD_REFERENCE, "bad reference")
            elif isinstance(arg, Decimal):
                if arg == Decimal(0):
                    return False
            elif isinstance(arg, str) and string_conversions.is_number(arg):
                if Decimal(arg) == Decimal(0):
                    return False
            elif isinstance(arg, str):
                if arg.lower() == "false":
                    return False
                if arg.lower() != "true":
                    return cell_error.CellError(
                        cell_error.CellErrorType.TYPE_ERROR, "Invalid string argument")
        return True

    def or_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        for arg in args:
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
            if isinstance(arg, unitialized_value.UninitializedValue):
                continue
            if isinstance(arg, bool) and not arg:
                continue
            elif isinstance(arg, Decimal):
                if arg == Decimal(0):
                    continue
            elif isinstance(arg, str) and string_conversions.is_number(arg):
                if Decimal(arg) == Decimal(0):
                    continue
            elif isinstance(arg, str):
                if arg.lower() == "false":
                    continue
                if arg.lower() != "true":
                    return cell_error.CellError(
                        cell_error.CellErrorType.TYPE_ERROR, "Invalid string argument")
            return True
        return False

    def not_func(self, args: List):
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        arg = args[0]
        if isinstance(arg, Function):
            arg = self.call_function(arg.name, arg.args)
        if isinstance(arg, unitialized_value.UninitializedValue):
            return True
        if isinstance(arg, bool) and not arg:
            return True
        elif isinstance(arg, Decimal):
            if arg == Decimal(0):
                return True
        elif isinstance(arg, str) and string_conversions.is_number(arg):
            if Decimal(arg) == Decimal(0):
                return True
        elif isinstance(arg, str):
            if arg.lower() == "false":
                return True
            if arg.lower() != "true":
                return cell_error.CellError(
                    cell_error.CellErrorType.TYPE_ERROR, "Invalid string argument")
        return False

    def xor_func(self, args: List):
        if len(args) == 0:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        true_cnt = 0
        for arg in args:
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
            if isinstance(arg, unitialized_value.UninitializedValue):
                continue
            if isinstance(arg, bool) and not arg:
                continue
            elif isinstance(arg, Decimal):
                if arg == Decimal(0):
                    continue
            elif isinstance(arg, str) and string_conversions.is_number(arg):
                if Decimal(arg) == Decimal(0):
                    continue
            elif isinstance(arg, str):
                if arg.lower() == "false":
                    continue
                if arg.lower() != "true":
                    return cell_error.CellError(
                        cell_error.CellErrorType.TYPE_ERROR, "Invalid string argument")
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
        if len(args) != 1:
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
        string_conversions.check_valid_location(args[0])
