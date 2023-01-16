import lark
import decimal
from lark.visitors import visit_children_decor
from sheets import cell_error
from sheets import workbook


# class CellRefFinder(lark.Visitor):
#     def __init__(self, current_sheet):
#         self.current_sheet = current_sheet
#         self.refs = []

#     def cell(self, tree):
#         self.refs.append(...)


# parser = lark.Lark.open('sheets/formulas.lark', start='formula')
# tree = parser.parse('=A1 + 3 * Sheet2!A2')
# v = CellRefFinder()
# v.visit(tree)


class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self):
        self.error = None
        self.sub_evaluator = None

    @visit_children_decor
    def add_expr(self, values):
        if values[1] == '+':
            return values[0] + values[2]
        elif values[1] == '-':
            return values[0] - values[2]
        else:
            assert False, 'Unexpected operator: ' + values[1]

    @visit_children_decor
    def mul_expr(self, values):
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
        if values[0] == "-":
            return -1 * values[1]

    @visit_children_decor
    def concat_expr(self, values):
        return values[0] + values[1]

    @visit_children_decor
    def cell(self, values, sheet=None):
        if len(values) > 1:
            sheet_name = values[0]
            sheet = workbook.spreadsheets[sheet_name]
            location = values[1]
        return sheet[location].value

    def parens(self, tree):
        self.sub_evaluator = FormulaEvaluator()
        return self.sub_evaluator.visit(tree.children[0])

    def number(self, tree):
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        return tree.children[0].value[1:-1]

    def error(self, tree):
        return self.error


# eval = FormulaEvaluator()
# parser = lark.Lark.open('sheets/formulas.lark', start='formula')
# tree = parser.parse("=Sheets1!A1")
# value = eval.visit(tree)
# print(f"value is {str(value)} with type {type(value)}")
