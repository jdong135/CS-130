class Sheet:
    def __init__(self, name: str):
        """_summary_

        Args:
            name (str): _description_
        """
        self.name = name
        self.extent_row = 0
        self.extent_col = 0
        self.cells = {}  # {(row, col): Cell}

    def add_cell(self, row: str, col: int, contents: str):
        """_summary_

        Args:
            row (str): _description_
            col (int): _description_
            contents (str): _description_
        """
        # Call evalContents from Workbook Class
        pass

    def modify_cell(self, row: str, col: int, new_contents: str):
        """_summary_

        Args:
            row (str): _description_
            col (int): _description_
            new_contents (str): _description_
        """
        pass

    def delete_cell(self, row: str, col: int):
        """_summary_

        Args:
            row (str): _description_
            col (int): _description_
        """
        pass

    def get_contents(self, row: str, col: int):
        """_summary_

        Args:
            row (str): _description_
            col (int): _description_
        """
        return self.cells[(row, col)].contents
