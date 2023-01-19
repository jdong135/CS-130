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

    def __hash__(self):
        return hash((self.sheet, self.location))

    def __init__(self, sheet: str, location: str, contents: str, value: str, type: CellType):
        """
        Initialize a cell object

        Args:
            sheet (str): Name of the sheet the cell is located in
            col (str): column of the cell - ranges from A-ZZZZ
            row (int): row of the cell - ranges from 1-9999
            contents (str): user input into the cell before evaluation
            value (str): evaluation of contents into an actual value
            type (CellType): what classification the cell is
        """
        self.sheet = sheet
        self.location = location
        self.contents = contents
        self.value = value
        self.type = type
        self.relies_on = set()  # which cells self calls
        self.dependents = set()  # which cells call self

    def set_fields(self, **kwargs) -> None:
        """
        Update specified fields of a cell object.
        """
        self.__dict__.update(kwargs)

    def set_value(self, value: str) -> None:
        """
        Set the value of a cell object.

        Args:
            value (str): value to set for the cell. 
        """
        self.value = value
