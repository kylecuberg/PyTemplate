import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# from context import basic_test


def main():
    # Main function

    try:
        True
    except Exception as E:
        print(type(E).__name__, __file__, E.__traceback__.tb_lineno, "\n", E)
    return None


if __name__ == "__main__":
    """[summary]"""
    main()
