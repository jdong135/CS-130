"""
Simple test for Github push. 
"""

import unittest


class GitAction(unittest.TestCase):
    def test_push(self):
        self.assertEqual("1", "1")


if __name__ == '__main__':
    unittest.main()
