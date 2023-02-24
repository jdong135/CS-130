from typing import List, Dict, Callable, Any
from decimal import Decimal
from sheets import cell_error, string_conversions, unitialized_value


class Function:
    def __init__(self, name, args):
        self.name = name
        self.args = args

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
                args.append(Function(sub_head, sub_args))         
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
            "XOR": self.xor_func
        }

    def call_function(self, func_name: str, args: List):
        try:
            return self.directory[func_name.strip().upper()](args)
        except KeyError:
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "Invalid Function name")

    def and_func(self, args: List):
        for arg in args:
            if isinstance(arg, Function):
                arg = self.call_function(arg.name, arg.args)
            if isinstance(arg, unitialized_value.UninitializedValue):
                return False
            if isinstance(arg, bool) and not arg:
                return False
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
    