import lark
import decimal
from lark.visitors import visit_children_decor

# class CellRefFinder(lark.Visitor):
#     def __init__(self, current_sheet):
#         self.current_sheet = current_sheetself.refs = []

#     def cell(self, tree):
#         self.refs.append(...)

# parser = lark.Lark.open('formulas.lark', start='formula')
# tree = parser.parse('=A1 + 3 * Sheet2!A2')
# v = CellRefFinder()
# v.visit(tree)


class FormulaEvaluator(lark.visitors.Interpreter):
    # def add_expr(self, tree):
    #     values = self.visit_children(tree)
    #     if values[1] == '+':
    #         return values[0] + values[2]
    #     elif values[1] == '-':
    #         return values[0] - values[2]
    #     else:
    #         assert False, 'Unexpecter operator: ' + values[1]

    @visit_children_decor
    def add_expr(self, values):
        if values[1] == '+':
            return values[0] + values[2]
        elif values[1] == '-':
            return values[0] - values[2]
        else:
            assert False, 'Unexpecter operator: ' + values[1]

    def number(self, tree):
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        return tree.children[0].value[1:-1]


parser = lark.Lark.open('formulas.lark', start='formula')
evaluator = FormulaEvaluator()
tree = parser.parse('=123.456 - 32.14')
value = evaluator.visit(tree)
print(f'value={value} type is {type(value)}')
