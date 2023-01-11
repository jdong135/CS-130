import lark


class Cell:
    def get_cell_type(self, contents):
        """_summary_

        Args:
            contents (_type_): _description_

        Returns:
            _type_: _description_
        """
        if contents[0] == "=":
            return "FORMULA"
        elif contents[0] == "'":
            return "STRING"
        else:
            return "NORMAL"

    def __init__(self, sheet: str, col: str, row: int, contents: str) -> None:
        """
        Initialize a cell object

        Args:
            sheet (str): Name of the sheet the cell is located in
            col (str): column of the cell - ranges from A-ZZZZ
            row (int): row of the cell - ranges from 1-9999
            contents (str): user input into the cell before evaluation
        """
        self.sheet = sheet
        self.col = col
        self.row = row
        self.contents = contents
        self.evaluation = None
        self.type = self.get_cell_type(contents)

    def set_evaluation(self, evaluation):
        """_summary_

        Args:
            evaluation (_type_): _description_
        """
        self.evaluation = evaluation
