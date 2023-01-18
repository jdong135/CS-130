

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
                    return cell_error.CellError(cell_error.CellErrorType.CIRCULAR_REFERENCE, "Cycle Detected In Topo Sort")
                    # update implementation when design decision is made
        else:
            result.append(v)
    return result[::-1]
