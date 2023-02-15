"""Representation of individual cell in spreadsheet."""

import enum
import uuid


class CellType(enum.Enum):
    """All possible types for a cell"""
    EMPTY = 1
    FORMULA = 2
    STRING = 3
    LITERAL_NUM = 4
    LITERAL_STRING = 5
    ERROR = 6


class Cell:
    """
    Cells are uniquely identified by the sheet they're in and their location.
    """

    def __eq__(self, obj):
        """Override equality for user defined class"""
        return obj.__dict__ == self.__dict__

    def __hash__(self):
        """Hash based off of the cell's sheet's id and the cell's location"""
        return hash((self.sheet.uuid, self.uuid))

    def __init__(self, sheet, location: str, contents: str, value: str, cell_type: CellType):
        """
        Initialize a cell object

        Args:
            sheet (Sheet): sheet object the cell is located in
            location (str): location of the cell in the form of `char` + `int`
            contents (str): user input into the cell before evaluation
            value (str): evaluation of contents into an actual value
            cell_type (CellType): what classification the cell is
        """
        self.sheet = sheet
        self.location = location
        self.contents = contents
        self.value = value
        self.cell_type = cell_type
        self.uuid = uuid.uuid1()

    def set_fields(self, **kwargs) -> None:
        """
        Update specified fields of a cell object.
        """
        self.__dict__.update(kwargs)
