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
