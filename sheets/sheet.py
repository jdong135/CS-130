"""Class that stores Cell objects that are all in the same spreadsheet."""
import uuid
from typing import Dict
import re
from sheets import string_conversions, cell


class Sheet:
    """
    Sheet class contains mapping of locations to Cell objects.
    Each sheet gets a unique ID for hashing purposes to check equality.
    Sheets track the extent of rows and columns.
    """

    def __eq__(self, obj):
        return isinstance(obj, Sheet) and obj.__dict__ == self.__dict__

    def __hash__(self):
        return hash(str(self.uuid))

    def __init__(self, name: str):
        """
        Initialize an empty sheet object with a given name. This sheet has a
        current extent of (0, 0) and empty dictionary of cell locations to Cell
        objects. 

        Args:
            name (str): name of the spreadsheet object.
        """
        self.name = name
        self.extent_row = 0
        self.extent_col = 0
        # {'A1': Cell} UPPERCASE location to cell
        self.cells: Dict[str, cell.Cell] = {}
        self.uuid: uuid.UUID = uuid.uuid1()
