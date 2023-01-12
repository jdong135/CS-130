"""
Test for implementation of Topological Sort algorithm. 
"""

import unittest
import os
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(__file__),
    os.pardir)
)
sys.path.append(PROJECT_ROOT)
from sheets import *

class GitAction(unittest.TestCase):
    def test_lecture_example(self):
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
        # iter_res = topo_sort(g)
        # for node in iter_res:
        #     print(node.value)


# if __name__ == '__main__':
#     unittest.main()
