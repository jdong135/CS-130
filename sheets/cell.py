import enum

class CellType(enum.Enum):
    FORMULA = 1
    STRING = 2
    LITERAL_NUM = 3


class Cell:
    def __init__(self, sheet: str, location: str, contents: str, value, type) -> None:
        """
        Initialize a cell object

        Args:
            sheet (str): Name of the sheet the cell is located in
            col (str): column of the cell - ranges from A-ZZZZ
            row (int): row of the cell - ranges from 1-9999
            contents (str): user input into the cell before evaluation
        """
        self.sheet = sheet
        self.location = location
        self.contents = contents
        self.value = value
        self.type = type
        self.neighbors = []

    def set_value(self, value):
        """_summary_

        Args:
            value (_type_): _description_
        """
        self.value = value
