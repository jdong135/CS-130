from typing import *

class Sheet:
    def __init__(self, name: str):
        """
        Initialize an empty sheet object with a given name. This sheet has a current extent of (0, 0) and empty dictionary of cell locations to Cell objects. 

        Args:
            name (str): name of the spreadsheet object.
        """
        self.name = name
        self.extent_row = 0
        self.extent_col = 0
        self.cells = {}  # {string: Cell}

    def str_to_tuple(self, location: str) -> Tuple[int, int]:
        """
        Take in a string location ranging from A1 to ZZZZ9999 and return the integer coordinates of the location.

        Args:
            location (str): string location on the sheet.

        Returns:
            Tuple[int, int]: numeric coordinates on the sheet equivalent to input location. 
        """
        chars = list(location)
        rows = 0
        cols = 0
        rowCnt = 0
        colCnt = 0
        for i in range(len(chars) - 1, -1, -1):
            if chars[i].isnumeric():
                rows += (10 ** rowCnt) * int(chars[i])
                rowCnt += 1
            else:
                cols += (26 ** colCnt) * (ord(chars[i]) - ord('A') + 1)
                colCnt += 1
        return (cols, rows)

    def check_valid_location(self, location: str) -> bool:
        """
        Determine if a cell location is witin the range of ZZZZ9999.

        Args:
            location (str): string location on the sheet.

        Returns:
            bool: if the input location is in-bounds.
        """
        col, row = self.str_to_tuple(location)
        if col > 475254 or row > 9999:
            return False
        return True

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
