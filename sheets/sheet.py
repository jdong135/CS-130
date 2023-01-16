class Sheet:
    def __init__(self, name: str):
        """_summary_

        Args:
            name (str): _description_
        """
        self.name = name
        self.extent_row = 0
        self.extent_col = 0
        self.cells = {}  # {string: Cell}

    def str_to_tuple(self, location: str):
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

    def check_valid_location(self, location: str):
        col, row = self.str_to_tuple(location)
        if col > 475254 or row > 9999:
            return False
        return True

