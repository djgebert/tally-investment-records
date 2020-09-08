import PyPDF2
import re

class InvestmentRecord():

    def __init__(self, filename: str) -> None:
        pass

    def populate(self, filename: str) -> None:

        # Ensure requested file has a pdf extension
        if not filename.endswith(".pdf"):
            raise ValueError("Input file must have extension .pdf")
        
        # Open the file
        f = open(filename, "rb")
        pdf_reader = PyPDF2.PdfFileReader(f)

        # Ensure it's a single page
        if pdf_reader.numPages > 1:
            raise Exception("Only single-page pdfs are accepted")

        # Get all the data we want
        text = pdf_reader.getPage(0).extractText()
        
        # Check it's an investment record we can handle
        # We can handle nabTrade Contract Notes
        # These will have a first line of text as follows
        if not text.startswith("WealthHub Securities Limited"):
            raise Exception("Only nabTrade Contract Notes are accepted")

        # Grab all the record details we're interested in
        trade_type = re.search(r"\n(.+) [Cc]onfirmation", text).group(1)
        settlement_date = re.search(r"Settlement date:\n(\d{2}/\d{2}/\d{4})", text).group(1)
        confirmation_number = re.search(r"Confirmation number:\n(\d+)", text).group(1)
        # details holds multiple named groups in one big match
        details = re.search(r"Consideration\n(?P<quantity>.+)\n(?P<code>.+)\n(?P<security_description>[^$]+)(?P<average_price>.+)\n(?P<consideration>.+)\n", text)
        quantity = details.group("quantity")
        code = details.group("code")
        average_price_per_share = details.group("average_price")

       