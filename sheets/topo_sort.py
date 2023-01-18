import enum
from sheets.cell import Cell
from sheets import cell_error


class DFSState(enum.Enum):
    ENTER = 1
    LEAVE = 2


def topo_sort(v: Cell) -> list:
    call_stack = [(v, DFSState.ENTER)]

    result = []
    visited = set()
    circular = False
    while call_stack:
        v, type = call_stack.pop()
        if type == DFSState.ENTER:
            visited.add((v.sheet, v.location))
            v.onStack = True
            call_stack.append((v, DFSState.LEAVE))
            for w in v.dependents:
                if (w.sheet, w.location) not in visited:
                    call_stack.append((w, DFSState.ENTER))
                # cycle detection
                if (w.sheet, w.location) in visited:
                    circular = True
                    # update implementation when design decision is made
        else:
            result.append(v)
    if not circular:
        return False, result[::-1]
    else:
        return True, result[::-1]
