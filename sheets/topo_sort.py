import enum


class Node:
    def __init__(self, value):
        self.onStack = False
        self.value = value
        self.neighbors = []


class DFSState(enum.Enum):
    ENTER = 1
    LEAVE = 2


def topo_sort(v: Node) -> list:
    call_stack = [(v, DFSState.ENTER)]

    result = []
    visited = set()

    while call_stack:
        v, type = call_stack.pop()
        if type == DFSState.ENTER:
            visited.add(v)
            v.onStack = True
            call_stack.append((v, DFSState.LEAVE))
            for w in v.neighbors:
                if w not in visited:
                    call_stack.append((w, DFSState.ENTER))
        else:
            result.append(v)
    return result[::-1]


# graph = Graph()
# a = Node(5)
# b = Node(7)
# c = Node(3)
# d = Node(11)
# e = Node(8)
# f = Node(2)
# g = Node(9)
# h = Node(10)
# a.neighbors = [d]
# b.neighbors = [d, e]
# c.neighbors = [e, h]
# d.neighbors = [f, g, h]
# e.neighbors = [g]
# graph.add_nodes([a, b, c, d, e, f, g, h])
# res = tarjan_recursive(b)
# for r in res:
#     print(r.value)

# iter_res = topo_sort(b)
# for node in iter_res:
#     print(node.value)

# a = Node(1)
# b = Node(2)
# c = Node(3)
# d = Node(4)
# e = Node(5)
# f = Node(6)
# g = Node(7)
# a.neighbors = [b, d]
# b.neighbors = [c, e]
# c.neighbors = [f]
# d.neighbors = [g]
# e.neighbors = [f]
# g.neighbors = [e, f]
# iter_res = topo_sort(g)
# for node in iter_res:
#     print(node.value)
