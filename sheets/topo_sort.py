"""This module topologically sorts a graph of Cells."""
import enum
from typing import Tuple, Dict, List
from sheets.cell import Cell


class DFSState(enum.Enum):
    """
    Enum used to determine if topo sort is visiting a Cell for the first time.
    """
    ENTER = 1
    LEAVE = 2


def topo_sort(cell: Cell, graph: Dict[Cell, List[Cell]]) -> Tuple[bool, List[Cell]]:
    """
    Perform a topological sort on all neighbors of the specified starting cell.

    Args:
        v (Cell): Cell to start the topological sort on.
        graph (Dict): Representation of a graph through an adjacency list. 

    Returns:
        Tuple[bool, list[Cell]]: Boolean indicating if the cell is part of a 
        cycle and the corresponding ordered list of topologically sorted cells.
    """
    call_stack = [(cell, DFSState.ENTER)]

    result = []
    visited = set()
    circular = False
    while call_stack:
        v, cell_state = call_stack.pop()
        if cell_state == DFSState.ENTER:
            visited.add((v.sheet, v.location))
            call_stack.append((v, DFSState.LEAVE))
            for w in graph[v]:
                if (w.sheet, w.location) not in visited:
                    call_stack.append((w, DFSState.ENTER))
                for c, s in call_stack:
                    if w.location == c.location and w.sheet == c.sheet and s == DFSState.LEAVE:
                        circular = True
                        break
        else:
            result.append(v)
    if not circular:
        return False, result[::-1]
    return True, result[::-1]
