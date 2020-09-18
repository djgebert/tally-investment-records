import PyPDF2
import re
import os
from typing import List

class InvestmentRecord():
    """Holds data for a single investment event, such as a sell or a buy.

    """

    def __init__(self, filename: str) -> None:
        pass

    def populate(self, filename: str) -> None:
        """Populate instance variables from a pdf file.

        Args:
            filename: The pdf file containing data for population.
            
        """
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
        self.trade_type = re.search(r"\n(.+) [Cc]onfirmation", text).group(1)
        self.settlement_date = re.search(r"Settlement date:\n(\d{2}/\d{2}/\d{4})", text).group(1)
        self.confirmation_number = re.search(r"Confirmation number:\n(\d+)", text).group(1)
        # details holds multiple named groups in one big match
        details = re.search(r"Consideration\n(?P<quantity>.+)\n(?P<code>.+)\n(?P<security_description>[^$]+)(?P<average_price>.+)\n(?P<consideration>.+)\n", text)
        self.quantity = details.group("quantity")
        self.code = details.group("code")
        self.average_price_per_share = details.group("average_price")

def get_contract_note_filenames(path: str) -> List[str]:
    """Return a list of filenames matching the contract note format (WH_ContractNote_....pdf).

    Args:
        path: The path to search. All subdirectories within the path are also searched.

    """
     
    contract_note_filenames = []
    for dirpath, dirnames, filenames in os.walk(path):
        for filename in filenames:
            if re.fullmatch(r"WH_ContractNote_.+\.pdf", filename):
                contract_note_filenames.append(os.path.join(dirpath, filename))
                
    return contract_note_filenames