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
            return values[0] + values[1]

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
                    new_empty_cell = cell.Cell(sheet, location, "", 0, cell.CellType.EMPTY)
                    sheet.cells[location] = new_empty_cell
                    new_empty_cell.add(self.calling_cell)
                    return 0 # FIX THIS
                return sheet.cells[location].value
            # =[col][row]
            else:
                if values[0].value not in self.sheet.cells:
                    # FIX THIS
                    location = values[0].value
                    new_empty_cell = cell.Cell(self.sheet, location, "", 0, cell.CellType.EMPTY)
                    self.sheet.cells[location] = new_empty_cell
                    new_empty_cell.neighbors.add(self.calling_cell)
                    return 0 # FIX THIS
                return self.sheet.cells[values[0].value].value

    def parens(self, tree):
        if not self.error:
            self.sub_evaluator = FormulaEvaluator(self.wb, self.sheet, self.calling_cell)
            return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        if not self.error:
            return decimal.Decimal(tree.children[0])

    def string(self, tree):
        if not self.error:
            return tree.children[0].value[1:-1]

    def error(self, tree):
        return self.error
