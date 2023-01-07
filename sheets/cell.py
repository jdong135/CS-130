import lark

class Cell:
    def get_literal(contents: str):
        """
        Assume the formula is valid. :)
        Assume the cell exists
        """
        contents = contents.rstrip()
        if contents[0] == "'":
            return contents[1:]
        if contents[0] == "=":
            lark.Lark.parse(contents[1:])

    def __init__(self, col: str, row: int, value: str):
        self.col = col
        self.row = row
        self.value = value