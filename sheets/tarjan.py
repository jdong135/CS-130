"""This module topologically sorts a graph of Cells."""
import enum
from typing import Tuple, Dict, List
from sheets.cell import Cell
import logging
logging.basicConfig(filename="logs/lark_module.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

class DFSState(enum.Enum):
    """
    Enum used to determine if topo sort is visiting a Cell for the first time.
    """
    NEW_RECURSIVE_CALL = 1
    PREV_RECURSIVE_CALL = 2

def tarjan(cell, graph: Dict[Cell, List[Cell]]) -> List[List[Cell]]:
    ret = []

    idx = 0
    
    stack = []
    stack_set = set()
    index_dict = {}
    lowlink_dict = {}

    call_stack = [(cell, DFSState.NEW_RECURSIVE_CALL, 0)]
    call_set = set([(cell, DFSState.NEW_RECURSIVE_CALL, 0)])

    while call_stack:
        v, recursion_state, recursion_idx = call_stack.pop()
        call_set.remove((v, recursion_state, recursion_idx))

        if recursion_state == DFSState.NEW_RECURSIVE_CALL:
            # Setting the depth index for v to the smallest unused index
            index_dict[v] = idx
            lowlink_dict[v] = idx
            idx += 1

            stack.append(v)
            stack_set.add(v)
        elif recursion_state == DFSState.PREV_RECURSIVE_CALL:
            recursive_vertex = graph[v][recursion_idx]
            lowlink_dict[v] = min(lowlink_dict[v], lowlink_dict[recursive_vertex])
            recursion_idx += 1
        while recursion_idx < len(graph[v]):
            if graph[v][recursion_idx] in index_dict:
                w = graph[v][recursion_idx]
                if w in stack_set:
                    lowlink_dict[v] = min(lowlink_dict[v], index_dict[w])
                recursion_idx += 1
            else:
                break
        if recursion_idx < len(graph[v]):
            w = graph[v][recursion_idx]
            call_stack.append((v, DFSState.PREV_RECURSIVE_CALL, recursion_idx))
            call_set.add((v, DFSState.PREV_RECURSIVE_CALL, recursion_idx))
            call_stack.append((w, DFSState.NEW_RECURSIVE_CALL, 0))
            call_set.add((w, DFSState.NEW_RECURSIVE_CALL, 0))
        if recursion_idx == len(graph[v]) and lowlink_dict[v] == index_dict[v]:
            scc = []
            w = None
            while w != v:
                w = stack.pop()
                stack_set.remove(w)
                scc.append(w)
            ret.append(scc)
    return ret      
