class Node:
    def __init__(self, value, index):
        self.onStack = False
        self.index = index
        self.value = value
        self.neighbors = []


class Graph:
    def __init__(self):
        self.vertices = []
        self.edges = set()  # (Node, Node)

    def add_nodes(self, list):
        for node in list:
            self.add_node(node)

    def add_node(self, node: Node):
        self.vertices.append(node)
        for neighbor in node.neighbors:
            self.edges.add((node, neighbor))

    def delete_node(self, node: Node):
        self.vertices.remove(node)
        for neighbor in node.neighbors:
            try:
                self.edges.remove((node, neighbor))
            except:
                print('failed deletion')


def tarjan(v: Node):
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

    # for v in graph.vertices:
    #     if not v.index:
    #         strongconnect(v)
    strongconnect(v)
    return result[::-1]


graph = Graph()
a = Node(5, None)
b = Node(7, None)
c = Node(3, None)
d = Node(11, None)
e = Node(8, None)
f = Node(2, None)
g = Node(9, None)
h = Node(10, None)
a.neighbors = [d]
b.neighbors = [d, e]
c.neighbors = [e, h]
d.neighbors = [f, g, h]
e.neighbors = [g]
graph.add_nodes([a, b, c, d, e, f, g, h])
res = tarjan(a)
assert res == [a, d, h, g, f]
