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
            "AND": self.and_func
        }

    def call_function(self, func_name: str, args: List):
        return self.directory[func_name.strip()](args)

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
