import os, sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from context import Project


def test(did_pass):
    """Print the result of a test."""
    linenum = sys._getframe(1).f_lineno  # Get the caller's line number.
    if did_pass:
        msg = f"Test at line {linenum} ok."
    else:
        msg = f"Test at line {linenum} FAILED."
    print(msg)


def main():
    return None


if __name__ == "__main__":
    """[summary]"""
    main()
