import pandas as pd
import xlsxwriter


def dataframe_to_excel(df, filename):
    """Improvement on df.csv - outputs dataframe to a table in an xlsx format.

    Args:
        df (dataframe): data to write to Excel
        filename (string): Print location & filename for xlsx file
    """
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    try:
        writer = pd.ExcelWriter(filename, engine="xlsxwriter")

        # Convert the dataframe to an XlsxWriter Excel object. Turn off the default
        # header and index and skip one row to allow us to insert a user defined
        # header.
        df.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False, index=False)

        # Get the xlsxwriter workbook and worksheet objects.
        # workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # Get the dimensions of the dataframe.
        (max_row, max_col) = df.shape

        # Create a list of column headers, to use in add_table().
        column_settings = []
        for header in df.columns:
            column_settings.append({"header": header})

        # Add the table.
        worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})

        # Make the columns wider for clarity.
        # worksheet.set_column(0, max_col - 1, 12)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
    except xlsxwriter.exceptions.FileCreateError as e:
        print("Could not create file", e)
    except Exception as E:
        print(type(E).__name__, __file__, E.__traceback__.tb_lineno, "\n", E)


def loop_replace(text_str="", replacements=[""]):
    """Does str.replace() for a list of strings.
    Args:
        text_str (str): text to remove from
        replacements (list, optional): list of strings to remove. Defaults to [""].
    Returns:
        string: text_str without the strings in replacements
    """

    if replacements != [""]:
        for i in replacements:
            text_str = text_str.replace(i, "")
    return text_str


def seconds_but_readable(seconds):
    seconds = seconds % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    return "%d:%02d:%02d" % (hour, minutes, seconds)
