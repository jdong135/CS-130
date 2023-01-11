import enum


class Node:
    def __init__(self, value):
        self.onStack = False
        self.index = None
        self.value = value
        self.neighbors = []
        self.lowlink = None


class DFSState(enum.Enum):
    ENTER = 1
    LEAVE = 2


def tarjan_iter(v: Node):
    result = []
    stack = [(v, DFSState.ENTER)]
    index = 0
    visited = set()

    while stack:
        node, state = stack.pop()
        if state == DFSState.LEAVE:
            result.append(node)
        else:
            visited.add(node)
            node.index = index
            node.lowlink = index
            index += 1
            stack.append((node, DFSState.LEAVE))
            for neighbor in node.neighbors:
                if neighbor not in visited:
                    stack.append((neighbor, DFSState.ENTER))
                    neighbor.onStack = True
                elif neighbor.onStack:
                    node.lowlink = min(node.lowlink, neighbor.lowlink)

    return result[::-1]


def tarjan_recursive(v: Node):
    result = []
    index = 0
    stack = []

    def strongconnect(v: Node):
        nonlocal index
        v.index = index
        v.lowlink = index
        index += 1
        stack.append(v)
        v.onStack = True

        for w in v.neighbors:
            if not w.index:
                strongconnect(w)
                v.lowlink = min(v.lowlink, w.lowlink)
            elif w.onStack:
                v.lowlink = min(v.lowlink, w.index)

        if v.lowlink == v.index:
            if stack:
                w = stack.pop()
                w.onStack = False
                result.append(w)
                while w != v:
                    if stack:
                        w = stack.pop()
                        w.onStack = False
                        result.append(w)
    strongconnect(v)
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

# iter_res = tarjan_iter(b)
# for node in iter_res:
#     print(node.value)

a = Node(1)
b = Node(2)
c = Node(3)
d = Node(4)
e = Node(5)
f = Node(6)
g = Node(7)
a.neighbors = [b, d]
b.neighbors = [c, e]
c.neighbors = [f]
d.neighbors = [g]
e.neighbors = [f]
g.neighbors = [e, f]
iter_res = tarjan_iter(g)
for node in iter_res:
    print(node.value)
