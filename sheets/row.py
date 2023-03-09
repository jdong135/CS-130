import decimal


class Row():
    def __init__(self, sort_cols, row_idx, cell_list) -> None:
        self.sort_cols = sort_cols
        self.row_idx = row_idx
        self.cell_list = cell_list
        # ex. cell_list = [1, 2, 3], sorting on col 3, sorting_cell_list = [3]
        self.sorting_cell_list = self.create_sorting_cell_list(sort_cols)

    def create_sorting_cell_list(self, sort_cols):
        res = []
        for i in sort_cols:
            val = self.cell_list[abs(i) - 1].value
            if i > 0:
                res.append(val)
            else:
                if isinstance(val, decimal.Decimal):
                    res.append(-1 * val)
                elif isinstance(val, bool):
                    res.append(not val)
                # add more types later
        return res

    def __eq__(self, obj):
        return isinstance(obj, Row) and self.cell_list == obj.cell_list

    def __lt__(self, obj):
        # for col in self.sort_cols:
        #     if col > 0:
        #         if self.cell_list[col - 1] < obj.cell_list[col - 1]:
        #             return True
        #         return False
        #     elif col < 0:
        #         if self.cell_list[abs(col) - 1] > obj.cell_list[abs(col) - 1]:
        #             return True
        #         return False
        # return False

        return self.sorting_cell_list < obj.sorting_cell_list
