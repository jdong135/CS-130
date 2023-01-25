import enum
from typing import *
from sheets.cell import Cell
from typing import Tuple


class DFSState(enum.Enum):
    ENTER = 1
    LEAVE = 2


def topo_sort(cell: Cell, graph) -> Tuple[bool, list[Cell]]:
    """
    Perform a topological sort on all neighbors of the specified starting cell.

    Args:
        v (Cell): Cell to start the topological sort on.

    Returns:
        Tuple[bool, list[Cell]]: Boolean indicating if the cell is part of a 
        cycle and the corresponding ordered list of topologically sorted cells.
    """
    call_stack = [(cell, DFSState.ENTER)]

    result = []
    visited = set()
    circular = False
    while call_stack:
        v, type = call_stack.pop()
        if type == DFSState.ENTER:
            visited.add((v.sheet, v.location))
            call_stack.append((v, DFSState.LEAVE))
            for w in graph[v]:
                if (w.sheet, w.location) not in visited:
                    call_stack.append((w, DFSState.ENTER))
                # cycle detection
                # if (w.sheet, w.location) in visited:
                #     circular = True
                #     # update implementation when design decision is made
                if (w.location, DFSState.ENTER) in call_stack or (w.location, DFSState.LEAVE) in call_stack:
                    circular = True
        else:
            result.append(v)
    if not circular:
        return False, result[::-1]
    else:
        return True, result[::-1]
