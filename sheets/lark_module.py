import lark
import decimal
from lark.visitors import visit_children_decor
from sheets import cell_error
from sheets import cell
from sheets import strip_module
from typing import Optional

# # importing module
# import logging

# # Create and configure logger
# logging.basicConfig(filename="logs/lark_module.log",
#                     format='%(asctime)s %(message)s',
#                     filemode='w')

# # Creating an object
# logger = logging.getLogger()

# # Setting the threshold of logger to DEBUG
# logger.setLevel(logging.DEBUG)


ERROR_MAP = {
    '#div/0!': cell_error.CellError(cell_error.CellErrorType.DIVIDE_BY_ZERO, "divide by zero"),
    '#ref!': cell_error.CellError(
        cell_error.CellErrorType.BAD_REFERENCE, "bad reference")}


class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, workbook, sheet, calling_cell):
        self.error = None
        self.sub_evaluator = None
        self.wb = workbook
        self.sheet = sheet
        self.calling_cell = calling_cell
        self.relies_on = set()

    # def __check_mapping(self, values):
    #     for i in [0, 2]:
    #         v = values[i]
    #         if v.lower() in ERROR_MAP:
    #             values[i] = ERROR_MAP[v.lower()]
    #     return values

    def __check_string_arithmetic(self, value) -> Optional[decimal.Decimal]:
        """
        Check if arithmetic is using values that are parsable to decimal values.
        If so, return their decimal form. Otherwise, set the FormulaEvaluator
        error to Type Error. 

        Args:
            value (any): value to check if parsable to a decimal

        Returns:
            Optional[decimal.Decimal]: decimal representation of input value
        """
        if type(value) == str:
            value = value.strip()
            if strip_module.is_number(value):
                value = decimal.Decimal(value)
                return value
            if len(value) > 0 or value:
                self.error = cell_error.CellError(
                    cell_error.CellErrorType.TYPE_ERROR,
                    "string arithmetic")
        else:
            return None

    def __check_for_error(self, values):
        """
        Check if input values is are instance of a cell error. If so, return the error with
        the highest priority. Otherwise, return False. 

        Args:
            values (List): array of input values for artithmetic evaluation

        Returns:
            bool or CellError: return self.error if an error is found. Otherwise, return false
        """
        errs = {}  # {int error enum : CellError object}
        for i in [0, 2]:
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
            self.error = errs[min(list(errs.keys()))]
            return self.error
        else:
            return False

    @visit_children_decor
    def add_expr(self, values):
        # logger.info(f"values are {values}")
        # values = self.__check_mapping(values)
        if self.__check_for_error(values):
            return self.error
        v0 = self.__check_string_arithmetic(values[0])
        v2 = self.__check_string_arithmetic(values[2])
        if self.error:
            return self.error
        if v0:
            values[0] = v0
        if v2:
            values[2] = v2
        if not values[0]:
            values[0] = decimal.Decimal(0)
        if not values[2]:
            values[2] = decimal.Decimal(0)
        if values[1] == '+':
            return values[0] + values[2]
        elif values[1] == '-':
            return values[0] - values[2]
        else:
            assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def mul_expr(self, values):
        # values = self.__check_mapping(values)
        if self.__check_for_error(values):
            return self.error
        v0 = self.__check_string_arithmetic(values[0])
        v2 = self.__check_string_arithmetic(values[2])
        if self.error:
            return self.error
        if v0:
            values[0] = v0
        if v2:
            values[2] = v2
        if not values[0]:
            values[0] = decimal.Decimal(0)
        if not values[2]:
            values[2] = decimal.Decimal(0)
        if values[1] == '*':
            res = values[0] * values[2]
            return abs(res) if res == 0 else res
        elif values[1] == '/':
            if values[2] == 0:
                self.error = cell_error.CellError(
                    cell_error.CellErrorType.DIVIDE_BY_ZERO, 'divide by zero')
                return self.error
            else:
                return values[0] / values[2]
        else:
            assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def unary_op(self, values):
        if values[0] == "-" and not values[1]:
            return decimal.Decimal(0)
        elif values[0] == "-":
            return -1 * values[1]

    @visit_children_decor
    def concat_expr(self, values):
        if not self.error:
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
        if not self.error:
            # =[sheet]![col][row]
            if len(values) > 1:
                # Make sure name has '' quotes if special characters
                sheet_name = values[0].value
                if sheet_name[0] == "'" and sheet_name[-1] == "'":  # check special
                    sheet_name = sheet_name.lower()[1:-1]
                else:
                    for ch in sheet_name:
                        if ch == " " or not (ch.isalnum() or ch == "_"):
                            self.error = cell_error.CellError(
                                cell_error.CellErrorType.PARSE_ERROR, 'special char requires single quotes')
                            return self.error
                    sheet_name = sheet_name.lower()
                if sheet_name not in self.wb.spreadsheets:
                    self.error = cell_error.CellError(
                        cell_error.CellErrorType.BAD_REFERENCE, 'sheet name not found')
                    return self.error
                sheet = self.wb.spreadsheets[sheet_name]
                location = values[1].value
                if not sheet.check_valid_location(location):
                    self.error = cell_error.CellError(
                        cell_error.CellErrorType.BAD_REFERENCE, 'invalid location')
                    return self.error
                # REFERENCE SELF CIRCULAR REFERENCE ERROR
                if self.calling_cell and self.calling_cell.sheet.lower() == sheet_name and self.calling_cell.location == location:
                    self.error = cell_error.CellError(
                        cell_error.CellErrorType.CIRCULAR_REFERENCE, 'cycle detected')
                    return self.error
                if location not in sheet.cells:
                    # FIX THIS
                    new_empty_cell = cell.Cell(
                        sheet.name, location, None, None, cell.CellType.EMPTY)
                    sheet.cells[location] = new_empty_cell
                    if self.calling_cell not in new_empty_cell.dependents:
                        new_empty_cell.dependents.add(self.calling_cell)
                    self.relies_on.add(new_empty_cell)
                    return decimal.Decimal(0)  # FIX THIS
                sheet.cells[location].dependents.add(self.calling_cell)
                self.relies_on.add(sheet.cells[location])
                # If you reference an empty cell, it's default value is none
                if not sheet.cells[location].value:
                    return decimal.Decimal(0)
                return sheet.cells[location].value
            # =[col][row]
            else:
                location = values[0].value
                if not self.sheet.check_valid_location(location):
                    self.error = cell_error.CellError(
                        cell_error.CellErrorType.BAD_REFERENCE, 'invalid location')
                    return self.error
                if self.calling_cell and self.calling_cell.location == location:
                    self.error = cell_error.CellError(
                        cell_error.CellErrorType.CIRCULAR_REFERENCE, 'cycle detected')
                    return self.error
                if values[0].value not in self.sheet.cells:
                    # FIX THIS
                    new_empty_cell = cell.Cell(
                        self.sheet.name, location, None, None, cell.CellType.EMPTY)
                    self.sheet.cells[location] = new_empty_cell
                    if self.calling_cell not in new_empty_cell.dependents:
                        new_empty_cell.dependents.add(self.calling_cell)
                    self.relies_on.add(new_empty_cell)
                    return decimal.Decimal(0)  # FIX THIS
                if self.calling_cell not in self.sheet.cells[location].dependents:
                    self.sheet.cells[location].dependents.add(
                        self.calling_cell)
                self.relies_on.add(self.sheet.cells[location])
                # If you reference an empty cell, it's default value is none
                if not self.sheet.cells[location].value:
                    return decimal.Decimal(0)
                return self.sheet.cells[values[0].value].value

    def parens(self, tree):
        if not self.error:
            self.sub_evaluator = FormulaEvaluator(
                self.wb, self.sheet, self.calling_cell)
            self.relies_on.update(self.sub_evaluator.relies_on)
            return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        if not self.error:
            number = decimal.Decimal(tree.children[0])
            if number == decimal.Decimal('NaN'):
                return "NaN"
            if number == decimal.Decimal('Infinity'):
                return "Infinity"
            return decimal.Decimal(tree.children[0])

    def string(self, tree):
        if not self.error:
            return tree.children[0].value[1:-1]

    def error(self, tree):
        val = tree.values[0].value.upper()

        if val == "#ERROR!":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.PARSE_ERROR, "input error")
            return cell_error.CellError(cell_error.CellErrorType.PARSE_ERROR, "input error")
        elif val == "#CIRCREF!":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
            return cell_error.CellError(cell_error.CellErrorType.CIRCULAR_REFERENCE, "input error")
        elif val == "#REF!":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.BAD_REFERENCE, "input error")
            return cell_error.CellError(cell_error.CellErrorType.BAD_REFERENCE, "input error")
        elif val == "#NAME?":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.BAD_NAME, "input error")
            return cell_error.CellError(cell_error.CellErrorType.BAD_NAME, "input error")
        elif val == "#VALUE!":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.TYPE_ERROR, "input error")
            return cell_error.CellError(cell_error.CellErrorType.TYPE_ERROR, "input error")
        elif val == "#DIV/0!":
            self.error = cell_error.CellError(
                cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")
            return cell_error.CellError(cell_error.CellErrorType.DIVIDE_BY_ZERO, "input error")

        return self.error


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
    eval = FormulaEvaluator(
        workbook, workbook.spreadsheets[sheetname.lower()], curr_cell)
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    try:
        tree = parser.parse(contents)
    except:
        eval.error = cell_error.CellError(
            cell_error.CellErrorType.PARSE_ERROR, "parse error")
        return eval, eval.error
    value = eval.visit(tree)
    return eval, value
    # except:
    #     curr_cell.value =
    #     eval.error = cell_error.CellError(
    #         cell_error.CellErrorType.BAD_REFERENCE, "bad reference")
    #     return eval, eval.error
