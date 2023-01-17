import lark
import decimal
from lark.visitors import visit_children_decor
from sheets import cell_error
from sheets import cell


class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, workbook, sheet, calling_cell):
        self.error = None
        self.sub_evaluator = None
        self.wb = workbook
        self.sheet = sheet
        self.calling_cell = calling_cell
        self.relies_on = set()

    @visit_children_decor
    def add_expr(self, values):
        if not self.error:
            if values[1] == '+':
                return values[0] + values[2]
            elif values[1] == '-':
                return values[0] - values[2]
            else:
                assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def mul_expr(self, values):
        if not self.error:
            if values[1] == '*':
                return values[0] * values[2]
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
        if not self.error:
            if values[0] == "-":
                return -1 * values[1]

    @visit_children_decor
    def concat_expr(self, values):
        if not self.error:
            return str(values[0]) + str(values[1])

    @visit_children_decor
    def cell(self, values):
        if not self.error:
            # =[sheet]![col][row]
            if len(values) > 1:
                sheet_name = values[0].value.lower()
                sheet = self.wb.spreadsheets[sheet_name]
                location = values[1].value
                if location not in sheet.cells:
                    # FIX THIS
                    new_empty_cell = cell.Cell(
                        sheet, location, "", "", cell.CellType.EMPTY)
                    sheet.cells[location] = new_empty_cell
                    if self.calling_cell not in new_empty_cell.dependents:
                        new_empty_cell.dependents.add(self.calling_cell)
                    self.relies_on.add(new_empty_cell)
                    return ""  # FIX THIS
                if self.calling_cell not in sheet.cells[location].dependents:
                    sheet.cells[location].dependents.add(self.calling_cell)
                self.relies_on.add(sheet.cells[location])
                return sheet.cells[location].value
            # =[col][row]
            else:
                location = values[0].value
                if values[0].value not in self.sheet.cells:
                    # FIX THIS

                    new_empty_cell = cell.Cell(
                        self.sheet, location, "", "", cell.CellType.EMPTY)
                    self.sheet.cells[location] = new_empty_cell
                    if self.calling_cell not in new_empty_cell.dependents:
                        new_empty_cell.dependents.add(self.calling_cell)
                    self.relies_on.add(new_empty_cell)
                    return ""  # FIX THIS
                if self.calling_cell not in self.sheet.cells[location].dependents:
                    self.sheet.cells[location].dependents.add(
                        self.calling_cell)
                self.relies_on.add(self.sheet.cells[location])
                return self.sheet.cells[values[0].value].value

    def parens(self, tree):
        if not self.error:
            self.sub_evaluator = FormulaEvaluator(
                self.wb, self.sheet, self.calling_cell)
            self.relies_on.update(self.sub_evaluator.relies_on)
            return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        if not self.error:
            return decimal.Decimal(tree.children[0])

    def string(self, tree):
        if not self.error:
            return tree.children[0].value[1:-1]

    def error(self, tree):
        return self.error


def evaluate_expr(workbook, curr_cell, sheetname, contents):
    eval = FormulaEvaluator(
        workbook, workbook.spreadsheets[sheetname.lower()], curr_cell)
    parser = lark.Lark.open('sheets/formulas.lark', start='formula')
    tree = parser.parse(contents)
    value = eval.visit(tree)
    return value
