"""Module containing functionality to parse spreadsheet formulas."""
import decimal
import re
from typing import Any, Union, Tuple, Callable
from functools import lru_cache
from typing import List
import lark
from contextlib import contextmanager, suppress
from lark.visitors import visit_children_decor
from lark.exceptions import UnexpectedInput
from sheets import cell_error, cell, string_conversions, unitialized_value, functions

import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class FormulaEvaluator(lark.visitors.Interpreter):
    """
    This class evaluates cell contents that begin with "=".
    It also tracks which cells the calling cell relies on.
    """

    def __init__(self, workbook, sheet, calling_cell):
        self.sub_evaluator = None
        self.wb = workbook
        self.sheet = sheet
        self.calling_cell = calling_cell
        self.calling_cell_relies_on = []
        self.parser = None
        self.__convert_literal_to_error = True

    @contextmanager
    def __ignore_error_literals(self):
        """
        Disable automatic conversions of error equivalent literals to CellError objects
        """
        self.__convert_literal_to_error = False
        yield
        self.__convert_literal_to_error = True

    def __check_sheet_name(self, sheet_name) -> Union[str, cell_error.CellError]:
        """
        Check if a sheet name is valid (comprised of valid characters and formatting).

        Args:
            sheet_name (str): sheet name to validate

        Returns:
            Union[str, cell_error.CellError]: lower-case form of the sheet name or an instance
            of a Parse Error if the sheet name is found to be invalid. 
        """
        if sheet_name[0] == " ":
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "invalid sheet name")
        if sheet_name[-1] == " ":
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "invalid sheet name")
        if re.match("\'[^']*\'", sheet_name):  # quoted sheet name
            sheet_name = sheet_name.lower()[1:-1]
        elif re.match("[A-Za-z_][A-Za-z0-9_]*", sheet_name):  # unquoted sheet name
            sheet_name = sheet_name.lower()
        else:
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "invalid sheet name")
        return sheet_name

    def __check_string_arithmetic(self, values, *args) -> Union[list, cell_error.CellError]:
        """
        Check a list of string values to confirm that they are all castable to the
        `decimal.Decimal` type. If they are, return the list of casted values. Otherwise, 
        return a Type Error indicating string arithmetic. 

        Args:
            values (list): input list of values generated by the parser
            args: indices of values array to check for string arithmetic

        Returns:
            Union[list, cell_error.CellError]: Either an updated list of decimal values with their
            corresponding operationg in string form, or an instance of a type error. 
        """
        for arg in args:
            assert isinstance(arg, int)
        assert len(args) in [1, 2]
        res = [decimal.Decimal(0), values[1], decimal.Decimal(0)] if len(
            args) == 2 else [values[0], decimal.Decimal(0)]
        for i in args:
            value = values[i]
            if isinstance(value, decimal.Decimal):
                res[i] = value
            elif isinstance(value, bool):
                res[i] = 1 if value else 0
            elif isinstance(value, unitialized_value.UninitializedValue):
                continue
            elif value and string_conversions.is_number(value):
                res[i] = decimal.Decimal(value)
            elif isinstance(value, str):
                return cell_error.CellError(
                    cell_error.CellErrorType.TYPE_ERROR, "string arithmetic")
        return res

    def __check_for_error(self, values, *args) -> Union[bool, cell_error.CellError]:
        """
        Check if input values is are instance of a cell error. If so, return the error with
        the highest priority. Otherwise, return False. 

        Args:
            values (List): array of input values for artithmetic evaluation
            args: indices of values array to check for cell error instances. If no arguments are
            provided, check the whole array.

        Returns:
            bool or CellError: return self.error if an error is found. Otherwise, return false
        """
        errs = {}  # {int error enum : CellError object}
        for arg in args:
            assert isinstance(arg, int)
        if len(args) == 0:
            args = range(len(values))
        for i in args:
            if isinstance(values[i], cell_error.CellError):
                match values[i].get_type():
                    case cell_error.CellErrorType.PARSE_ERROR:
                        errs[cell_error.CellErrorType.PARSE_ERROR.value] = cell_error.CellError(
                            cell_error.CellErrorType.PARSE_ERROR, "parsing error")
                    case cell_error.CellErrorType.CIRCULAR_REFERENCE:
                        errs[cell_error.CellErrorType.CIRCULAR_REFERENCE.value] = \
                            cell_error.CellError(
                            cell_error.CellErrorType.CIRCULAR_REFERENCE, "circular reference")
                    case cell_error.CellErrorType.BAD_REFERENCE:
                        errs[cell_error.CellErrorType.BAD_REFERENCE.value] = cell_error.CellError(
                            cell_error.CellErrorType.BAD_REFERENCE, "bad reference")
                    case cell_error.CellErrorType.BAD_NAME:
                        errs[cell_error.CellErrorType.BAD_NAME.value] = cell_error.CellError(
                            cell_error.CellErrorType.BAD_NAME, "bad name")
                    case cell_error.CellErrorType.TYPE_ERROR:
                        errs[cell_error.CellErrorType.TYPE_ERROR.value] = cell_error.CellError(
                            cell_error.CellErrorType.TYPE_ERROR, "invalid operation")
                    case cell_error.CellErrorType.DIVIDE_BY_ZERO:
                        errs[cell_error.CellErrorType.DIVIDE_BY_ZERO.value] = cell_error.CellError(
                            cell_error.CellErrorType.DIVIDE_BY_ZERO, "divide by zero")
        if len(errs) > 0:
            return errs[min(list(errs.keys()))]
        return False

    def __evaluate_function_arguments(self, args_list: List[str]) -> Tuple[List[Any], Union[bool, cell_error.CellError]]:
        """
        Evaluate all arguments of a function and any composed function.

        Args:
            args_list (List[str]): List of string arguments

        Returns:
            Tuple[List[Any], bool | CellError]: Tuple containing a list of evaluated arguments
            and if a CellError is present in the evaluation. If an error is present, the error 
            with highest priority is returned. If no CellError is detected, this field is False.
        """
        error_found = False
        for i in range(len(args_list)):
            arg = args_list[i]
            if isinstance(arg, functions.Function):
                arg.args, temp_err = self.__evaluate_function_arguments(
                    arg.args)
                # We only overwrite with new errors if we haven't already found one
                # We don't want to make error_found false again if we've already found an error
                if not error_found:
                    error_found = temp_err
            else:
                try:
                    tree = get_tree(self.parser, "=" + arg)
                    arg = self.visit(tree)
                except lark.exceptions.UnexpectedInput:
                    continue
            args_list[i] = arg
        # If no error has been found, check each arg for cell error instance
        if not error_found:
            error_found = self.__check_for_error(args_list)
        return args_list, error_found

    def __evaluate_and_solve_arg(self, func, i) -> Tuple[Any, Union[bool, cell_error.CellError]]:
        [val], error_found = self.__evaluate_function_arguments([func.args[i]])
        if not error_found and isinstance(val, functions.Function):
            val = self.wb.function_directory.call_function(
                val.name, val.args)
        return val, error_found

    def __bool_cmpr(self, left: Any, right: Any,
                    operand: Callable[[str, str], bool], string_op: str) -> bool:
        # booleans > strings > numbers
        if type(left) == type(right):
            return operand(left, right)
        if isinstance(left, bool):
            return True if string_op in ['>', '>='] else False
        if isinstance(left, str):
            if isinstance(right, decimal.Decimal):
                return True if string_op in ['>', '>='] else False
            return False if string_op in ['>', '>='] else True
        if isinstance(left, decimal.Decimal):
            return False if string_op in ['>', '>='] else True

    @visit_children_decor
    def add_expr(self, values):
        potential_error = self.__check_for_error(values, 0, 2)
        if potential_error:
            return potential_error
        updated_values = self.__check_string_arithmetic(values, 0, 2)
        if isinstance(updated_values, cell_error.CellError):
            return updated_values
        if updated_values[1] == '+':
            return updated_values[0] + updated_values[2]
        if updated_values[1] == '-':
            return updated_values[0] - updated_values[2]
        assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def mul_expr(self, values):
        potential_error = self.__check_for_error(values, 0, 2)
        if potential_error:
            return potential_error
        updated_values = self.__check_string_arithmetic(values, 0, 2)
        if isinstance(updated_values, cell_error.CellError):
            return updated_values
        if updated_values[1] == '*':
            res = updated_values[0] * updated_values[2]
            return abs(res) if res == 0 else res
        if updated_values[1] == '/':
            if updated_values[2] == 0:
                return cell_error.CellError(
                    cell_error.CellErrorType.DIVIDE_BY_ZERO, 'divide by zero')
            return updated_values[0] / updated_values[2]
        assert False, 'Unexpected operator: ' + updated_values[1]

    @visit_children_decor
    def unary_op(self, values):
        logger.info(values)
        potential_error = self.__check_for_error(values, 1)
        if potential_error:
            return potential_error
        updated_values = self.__check_string_arithmetic(values, 1)
        if isinstance(updated_values, cell_error.CellError):
            return updated_values
        if updated_values[0] == "+":
            return updated_values[1]
        if updated_values[0] == "-":
            return -1 * updated_values[1]
        return None

    @visit_children_decor
    def concat_expr(self, values):
        potential_error = self.__check_for_error(values, 0, 1)
        if potential_error:
            return potential_error
        if isinstance(values[1], decimal.Decimal) and values[1] == decimal.Decimal(0):
            values[1] = "0"
        elif isinstance(values[1], bool):
            values[1] = "TRUE" if values[1] else "FALSE"
        elif not values[1] or isinstance(values[1], unitialized_value.UninitializedValue):
            values[1] = ""
        if isinstance(values[0], decimal.Decimal) and values[0] == decimal.Decimal(0):
            values[0] = "0"
        elif isinstance(values[0], bool):
            values[0] = "TRUE" if values[0] else "FALSE"
        elif not values[0] or isinstance(values[0], unitialized_value.UninitializedValue):
            values[0] = ""
        res = str(values[0]) + str(values[1])
        return res

    @visit_children_decor
    def bool_expr(self, values):
        logger.info(values)
        potential_error = self.__check_for_error(values, 0, 2)
        if potential_error:
            return potential_error
        left, operator, right = values
        ###
        if isinstance(left, str):
            left = left.lower()
            if left == "true":
                left = True
            elif left == "false":
                left = False
        ##
        elif isinstance(left, unitialized_value.UninitializedValue):
            if isinstance(right, str):
                left = ""
            elif isinstance(right, decimal.Decimal):
                left = decimal.Decimal(0)
            elif isinstance(right, bool):
                left = False
        # ASK DONNIE
        if isinstance(right, str):
            right = right.lower()
            if right == "true":
                right = True
            elif right == "false":
                right = False
        ###
        elif isinstance(right, unitialized_value.UninitializedValue):
            if isinstance(left, str):
                right = ""
            elif isinstance(left, decimal.Decimal):
                right = decimal.Decimal(0)
            elif isinstance(left, bool):
                right = False
        if operator == "=" or operator == "==":
            return left == right
        elif operator == "<>" or operator == "!=":
            return left != right
        elif operator == ">":
            return self.__bool_cmpr(left, right, lambda x, y: x > y, '>')
        elif operator == ">=":
            return self.__bool_cmpr(left, right, lambda x, y: x >= y, '>=')
        elif operator == "<":
            return self.__bool_cmpr(left, right, lambda x, y: x < y, '<')
        elif operator == "<=":
            return self.__bool_cmpr(left, right, lambda x, y: x <= y, '<=')
        assert False, 'Unexpected operator: ' + operator

    def function(self, values):        
        name = values.children[0].strip().upper()
        func = functions.Function(name, [], False)
        if name not in self.wb.function_directory.get_function_keys():
            return cell_error.CellError(cell_error.CellErrorType.BAD_NAME, "Invalid function name")
        if name in self.wb.function_directory.lazy_functions:
            func.lazy_eval = True
        if not func.lazy_eval:
            for child in values.children[1:]:
                func.args.append(self.visit(child))
        else:  # lazy evaluation  
            if func.name == "IF":
                logger.info(values.children[1:])
                if len(values.children[1:]) not in [2, 3]:
                    return cell_error.CellError(
                        cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")                
                if string_conversions.check_for_true_arg(self.visit(values.children[1])):
                    func.args.append(self.visit(values.children[2]))
                elif len(values.children) > 2:
                    func.args.append(values.children[3])
                else:
                    func.args.append(False)
            elif func.name == "IFERROR":
                if len(values.children[1:]) not in [1, 2]:
                    return cell_error.CellError(
                        cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
                with self.__ignore_error_literals():
                    res = self.visit(values.children[1])
                if not isinstance(res, cell_error.CellError):
                    func.args.append(res)
                elif len(values.children[1:]) > 1:
                    func.args.append(self.visit(values.children[2]))
                else:
                    func.args.append("")
            # if error_found and func.name != "ISERROR":
            #     return error_found
            # if func.name == "ISERROR":
            #     if len(func.args) != 1:
            #         return cell_error.CellError(
            #             cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
            #     if isinstance(func.args[0], str):
            #         try:
            #             get_tree(self.parser, "=" + func.args[0])
            #         except lark.exceptions.UnexpectedInput:
            #             return cell_error.CellError(
            #                 cell_error.CellErrorType.PARSE_ERROR, "Input cannot be parsed")
            # elif func.name == "INDIRECT":
            #     if len(func.args) != 1:
            #         return cell_error.CellError(
            #             cell_error.CellErrorType.TYPE_ERROR, "Invalid argument count")
            #     if not isinstance(func.args[0], str) \
            #         or not string_conversions.check_valid_location(func.args[0]) \
            #             or isinstance(func.args[0], unitialized_value.UninitializedValue):
            #         return cell_error.CellError(
            #             cell_error.CellErrorType.BAD_REFERENCE, "Bad reference")
            #     func.args[0] = self.visit(
            #         get_tree(self.parser, "=" + func.args[0]))
        return self.wb.function_directory.call_function(func)

    @visit_children_decor
    def cell(self, values):
        logger.info(f"cell ref in {values}")
        # handle different input formats for value
        if len(values) > 1:  # =[sheet]![col][row]
            sheet_name = self.__check_sheet_name(values[0].value)
            if isinstance(sheet_name, cell_error.CellError):
                return sheet_name
            location = values[1].value.upper().replace("$", "")
        else:  # = [col][row]
            sheet_name = self.sheet.name.lower()
            location = values[0].value.upper().replace("$", "")

        # check for invalid references or errors
        if sheet_name not in self.wb.spreadsheets:
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_REFERENCE, "sheet name not found")
        sheet = self.wb.spreadsheets[sheet_name]
        if not string_conversions.check_valid_location(location):
            return cell_error.CellError(cell_error.CellErrorType.BAD_REFERENCE, "invalid location")
        if self.calling_cell and self.calling_cell.sheet.name.lower() == sheet_name \
                and self.calling_cell.location == location:
            return cell_error.CellError(
                cell_error.CellErrorType.CIRCULAR_REFERENCE, "circular reference")

        # no cell in this location yet
        if location not in sheet.cells:
            new_empty_cell = cell.Cell(
                sheet, location, None, None, cell.CellType.EMPTY)
            sheet.cells[location] = new_empty_cell
            self.wb.adjacency_list[new_empty_cell] = [self.calling_cell]
            self.calling_cell_relies_on.append(new_empty_cell)
            return unitialized_value.UninitializedValue()

        # add calling_cell to the neighbors of Cell from values argument
        if self.calling_cell not in self.wb.adjacency_list[sheet.cells[location]]:
            self.wb.adjacency_list[sheet.cells[location]].append(
                self.calling_cell)
        self.calling_cell_relies_on.append(sheet.cells[location])
        # cell exists at location but is empty
        if not sheet.cells[location].value and not isinstance(sheet.cells[location].value, bool):
            ret_val = unitialized_value.UninitializedValue()
        else:
            ret_val = sheet.cells[location].value
        return ret_val

    def parens(self, tree):
        self.sub_evaluator = FormulaEvaluator(
            self.wb, self.sheet, self.calling_cell)
        self.calling_cell_relies_on.extend(
            self.sub_evaluator.calling_cell_relies_on)
        return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        number = decimal.Decimal(tree.children[0])
        if number == decimal.Decimal('NaN'):
            return "NaN"
        if number == decimal.Decimal('Infinity'):
            return "Infinity"
        return decimal.Decimal(string_conversions.strip_zeros(tree.children[0]))

    def string(self, tree):
        value = tree.children[0].value[1:-1]
        if self.__convert_literal_to_error:
            potential_error = string_conversions.str_to_error(value)
            if potential_error:
                return potential_error
        return value

    def boolean(self, tree):
        value = tree.children[0].value
        if string_conversions.is_true_expr(value):
            return True
        return False

    def error(self, tree):
        match tree.children[0].upper():
            case "#ERROR!":
                return cell_error.CellError(
                    cell_error.CellErrorType.PARSE_ERROR, "input error")
            case "#CIRCREF!":
                return cell_error.CellError(
                    cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
            case "#REF!":
                return cell_error.CellError(
                    cell_error.CellErrorType.BAD_REFERENCE, "input error")
            case "#NAME?":
                return cell_error.CellError(
                    cell_error.CellErrorType.BAD_NAME, "input error")
            case "#VALUE!":
                return cell_error.CellError(
                    cell_error.CellErrorType.TYPE_ERROR, "input error")
            case "#DIV/0!":
                return cell_error.CellError(
                    cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")


@lru_cache(maxsize=None)
def open_grammar() -> lark.Lark:
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    return parser


@lru_cache(maxsize=None)
def get_tree(parser: lark.Lark, contents: str) -> lark.ParseTree:
    return parser.parse(contents)


def evaluate_expr(workbook, curr_cell, sheetname: str,
                  contents: str) -> tuple[FormulaEvaluator, Any]:
    """
    Evaluate a provided expression using the lark formula parser and evaluator.

    Args:
        workbook (Workbook): a workbook object
        curr_cell (Cell): cell to evaluate within
        sheetname (str): name of the sheet we are working in
        contents (str): contents of the cell to parse

    Returns:
        FormulaEvaluator, Any: Evaluator object and provided value 
    """
    if sheetname.lower() not in workbook.spreadsheets:
        return None, cell_error.CellError(
            cell_error.CellErrorType.BAD_REFERENCE, "bad reference")
    sheet = workbook.spreadsheets[sheetname.lower()]
    evaluator = FormulaEvaluator(workbook, sheet, curr_cell)
    parser = open_grammar()
    evaluator.parser = parser
    logger.info(f"contents: {contents}")
    try:
        tree = get_tree(parser, contents)
    except UnexpectedInput:
        logger.info("here")
        return evaluator, cell_error.CellError(
            cell_error.CellErrorType.PARSE_ERROR, "parse error")
    logger.info(f"tree: {tree}")
    value = evaluator.visit(tree)
    return evaluator, value
