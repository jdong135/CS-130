"""
Implementation of UnitializedValue object. 
"""


class UninitializedValue():
    """
    Placeholder value for any cell with an unitialized value. This removes
    prevents us from needing to equate `None` contents with 0, since we have
    a unique value object to represent this. This enables us to distinguish
    between cells that evaluate to a value of 0 and cells with `None` contents.
    """

    def __init__(self):
        return

    def __eq__(self, other):
        return isinstance(other, UninitializedValue)

    def __ne__(self, other):
        return not isinstance(other, UninitializedValue)

    def __str__(self):
        return ""

class EmptySortCell():
    """
    Placeholder value for any empty cell during sort.
    """

    def __init__(self):
        return

    def __eq__(self, other):
        return isinstance(other, EmptySortCell)

    def __ne__(self, other):
        return not isinstance(other, EmptySortCell)

    def __str__(self):
        return ""
