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
    ENTER = 1
    LEAVE = 2

def tarjan(cel, graph: Dict[Cell, List[Cell]]) -> List[List[Cell]]:
    index = 0
    call_stack = []
    call_set = set()

    stack = []
    stack_set = set()
    index_dict = {}
    lowlink_dict = {}
    scc = []


   
    call_stack.append((cel, 0))
    call_set.add((cel, 0))
    while call_stack:
        c, state = call_stack.pop()
        call_set.remove((c, state))
        if state == 0:
            index_dict[c] = index
            lowlink_dict[c] = index
            index += 1
            stack.append(c)
            stack_set.add(c)
        elif state > 0:
            p = graph[c][state - 1]
            lowlink_dict[c] = min(lowlink_dict[c], lowlink_dict[p])
        while state < len(graph[c]) and graph[c][state] in index_dict:
            w = graph[c][state]
            if w in stack_set:
                lowlink_dict[c] = min(lowlink_dict[c], lowlink_dict[w])
            state += 1
        if state < len(graph[c]):
            w = graph[c][state]
            call_stack.append((c, state + 1))
            call_set.add((c, state + 1))
            call_stack.append((w, 0))
            call_set.add((w, 0))
            continue
        if lowlink_dict[c] == index_dict[c]:
            cc = []
            while True:
                w = stack.pop()
                stack_set.remove(w)
                cc.append(w)
                if w == c:
                    break
            scc.append(cc)

    return scc








