import os
import pandas as pd
import xlsxwriter
import smtplib
import ssl
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


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
        print(E, type(E).__name__, __file__, E.__traceback__.tb_lineno)


class the_email:
    """Email object. Currently configured for gmail, prompts for password."""

    def __init__(self):
        """Creation of object

        Returns:
            email: email-ready object
        """
        self.port = 465  # For SSL
        self.context = ssl.create_default_context()
        self.sender_email = self.get_input("email")
        self.sender_password = self.get_input("password")
        return None

    def get_input(self, str_of_item):
        """Get inputs to email

        Args:
            str_of_item (string): string to ask user for

        Returns:
            string: user-entered string
        """
        return input(f"Type {str_of_item} and press enter: ")

    def email_out(self, subject, body, to, cc="", bcc="", attachment=""):
        """Sends email with given parameters

        Args:
            subject (str): _description_
            body (str): string/html of what the body of the email should say
            to (str): To for email
            cc (str, optional): CC  for email. Defaults to "".
            bcc (str, optional): BCC for email. Defaults to "".
            attachment (str, optional): Filepath/filename of any attachments to include. Defaults to "".
        """

        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        message["From"] = self.sender_email
        message["To"] = to
        message["Cc"] = cc
        message["Bcc"] = bcc
        message.attach(MIMEText(body, "html"))
        if attachment != "":
            with open(attachment, "rb") as f:
                message.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="xlsx",
                    filename=attachment,
                )
                attachment_filename = os.path.basename(attachment)
                mime_type, _ = mimetypes.guess_type(attachment)
                mime_type, mime_subtype = mime_type.split("/", 1)
                message.add_attachment(
                    f.read(),
                    maintype=mime_type,
                    subtype=mime_subtype,
                    filename=attachment_filename,
                )

        with smtplib.SMTP_SSL(
            "smtp.gmail.com", self.port, context=self.context
        ) as server:
            server.login(self.sender_email, self.sender_password)
            server.send_message(message)


def loop_remove(text_str="", replacements=[""]):
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
