import lark
import decimal
from lark.visitors import visit_children_decor
import cell_error


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
    # def add_expr(self, tree):
    #     values = self.visit_children(tree)
    #     print(values)
    #     if values[1] == '+':
    #         return values[0] + values[2]
    #     elif values[1] == '-':
    #         return values[0] - values[2]
    #     else:
    #         assert False, 'Unexpecter operator: ' + values[1]
    def __init__(self):
        self.error = None

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
    def concat_expr(self, values):
        return values[0] + values[1]

    def parens(self, values):
        return

    def number(self, tree):
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        return tree.children[0].value[1:-1]

    def error(self, tree):
        return self.error


parser = lark.Lark.open('sheets/formulas.lark', start='formula')
evaluator = FormulaEvaluator()
tree = parser.parse('=a1 & b1')
value = evaluator.visit(tree)
print(f'value={value} type is {type(value)}')
