class UninitializedValue():
    def __init__(self):
        return

    def __eq__(self, other):
        return isinstance(other, UninitializedValue)

    def __ne__(self, other):
        return not isinstance(other, UninitializedValue)
