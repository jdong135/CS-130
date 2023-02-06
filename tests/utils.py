"""
Utility functions and helpers for test files.
"""

import sys
import io


def store_stdout():
    sys_out = sys.stdout
    new_stdo = io.StringIO()
    sys.stdout = new_stdo
    return new_stdo, sys_out


def restore_stdout(new_stdo, sys_out):
    output = new_stdo.getvalue()
    sys.stdout = sys_out
    return output
