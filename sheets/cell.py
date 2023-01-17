import enum


class CellType(enum.Enum):
    EMPTY = 1
    FORMULA = 2
    STRING = 3
    LITERAL_NUM = 4
    LITERAL_STRING = 5


class Cell:
    def __eq__(self, obj):
        return isinstance(obj, Cell) and obj.__dict__ == self.__dict__

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
        self.neighbors = set()

    def set_value(self, value):
        """_summary_

        Args:
            value (_type_): _description_
        """
        self.value = value
