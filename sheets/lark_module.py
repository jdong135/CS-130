import lark
import decimal
from lark.visitors import visit_children_decor
from sheets import cell_error
from sheets import cell
from sheets import strip_module
from typing import Optional, Any, Union

import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, workbook, sheet, calling_cell):
        self.sub_evaluator = None
        self.wb = workbook
        self.sheet = sheet
        self.calling_cell = calling_cell

    def __check_string_arithmetic(self, values) -> Union[list, cell_error.CellError]:
        res = [decimal.Decimal(0), values[1], decimal.Decimal(0)]
        for i in [0, 2]: 
            value = values[i]
            if not value:
                continue
            elif isinstance(value, decimal.Decimal):
                res[i] = value
            elif value and strip_module.is_number(value):
                res[i] = decimal.Decimal(value)
            else:
                return cell_error.CellError(cell_error.CellErrorType.TYPE_ERROR, "string arithmetic")
        return res

    def __check_for_error(self, values, *args) -> Union[bool, cell_error.CellError]:
        """
        Check if input values is are instance of a cell error. If so, return the error with
        the highest priority. Otherwise, return False. 

        Args:
            values (List): array of input values for artithmetic evaluation
            args: indices of values array to check for cell error instances

        Returns:
            bool or CellError: return self.error if an error is found. Otherwise, return false
        """
        errs = {}  # {int error enum : CellError object}
        for arg in args:
            assert type(arg) == int
        for i in args:
            if isinstance(values[i], cell_error.CellError):
                match values[i].get_type():
                    case cell_error.CellErrorType.PARSE_ERROR:
                        errs[cell_error.CellErrorType.PARSE_ERROR.value] = cell_error.CellError(
                            cell_error.CellErrorType.PARSE_ERROR, "parsing error")
                    case cell_error.CellErrorType.CIRCULAR_REFERENCE:
                        errs[cell_error.CellErrorType.CIRCULAR_REFERENCE.value] = cell_error.CellError(
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
        else:
            return False

    @visit_children_decor
    def add_expr(self, values):
        potential_error = self.__check_for_error(values, 0, 2)
        if potential_error:
            return potential_error
        updated_values = self.__check_string_arithmetic(values)
        if isinstance(updated_values, cell_error.CellError):
            return updated_values
        if updated_values[1] == '+':
            return updated_values[0] + updated_values[2]
        elif updated_values[1] == '-':
            return updated_values[0] - updated_values[2]
        else:
            assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def mul_expr(self, values):
        potential_error = self.__check_for_error(values, 0, 2)
        if potential_error:
            return potential_error
        updated_values = self.__check_string_arithmetic(values)
        if isinstance(updated_values, cell_error.CellError):
            return updated_values
        if values[1] == '*':
            res = values[0] * values[2]
            return abs(res) if res == 0 else res
        elif values[1] == '/':
            if values[2] == 0:
                return cell_error.CellError(
                    cell_error.CellErrorType.DIVIDE_BY_ZERO, 'divide by zero')
            else:
                return values[0] / values[2]
        else:
            assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def unary_op(self, values):
        potential_error = self.__check_for_error(values, 1)
        if potential_error:
            return potential_error
        if values[0] == "-" and not values[1]:
            return decimal.Decimal(0)
        elif values[0] == "-":
            return -1 * values[1]

    @visit_children_decor
    def concat_expr(self, values):
        if not self.cell_error:
            self.sub_evaluator = FormulaEvaluator(
                self.wb, self.sheet, self.calling_cell)
            self.relies_on.update(self.sub_evaluator.relies_on)
            if not values[1]:
                values[1] = ""
            if not values[0]:
                values[0] = ""
            return str(values[0]) + str(values[1])

    @visit_children_decor
    def cell(self, values):
        if not self.cell_error:
            pass


    def parens(self, tree):
        self.sub_evaluator = FormulaEvaluator(
            self.wb, self.sheet, self.calling_cell)
        return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        number = decimal.Decimal(tree.children[0])
        if number == decimal.Decimal('NaN'):
            return "NaN"
        if number == decimal.Decimal('Infinity'):
            return "Infinity"
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        return tree.children[0].value[1:-1]

    def error(self, tree):
        val = tree.children[0].upper()

        if val == "#ERROR!":
            return cell_error.CellError(
                cell_error.CellErrorType.PARSE_ERROR, "input error")
        elif val == "#CIRCREF!":
            return cell_error.CellError(
                cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
        elif val == "#REF!":
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_REFERENCE, "input error")
        elif val == "#NAME?":
            return cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "input error")
        elif val == "#VALUE!":
            return cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "input error")
        elif val == "#DIV/0!":
            return cell_error.CellError(
                cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")


def evaluate_expr(workbook, curr_cell, sheetname, contents):
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
    try:
        eval = FormulaEvaluator(
            workbook, workbook.spreadsheets[sheetname.lower()], curr_cell)
    except KeyError:
        return None, cell_error.CellError(
            cell_error.CellErrorType.BAD_REFERENCE, "bad reference")
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    try:
        tree = parser.parse(contents)
    except:
        eval.cell_error = cell_error.CellError(
            cell_error.CellErrorType.PARSE_ERROR, "parse error")
        return eval, eval.cell_error
    value = eval.visit(tree)
    return eval, value
