import decimal


class Row():
    def __init__(self, sort_cols, row_idx, cell_list) -> None:
        self.sort_cols = sort_cols
        self.row_idx = row_idx
        self.cell_list = cell_list
        self.rev_list = self.rev_needed(sort_cols)

    def rev_needed(self, sort_cols):
        res = []
        for c in sort_cols:
            if c > 0:
                res.append(False)
            else:
                res.append(True)
        return res

    def __eq__(self, obj):
        return isinstance(obj, Row) and self.cell_list == obj.cell_list
