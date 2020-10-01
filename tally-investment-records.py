import PyPDF2
import re
import os
from typing import List
import sys
from openpyxl import Workbook
# import defusedxml
from collections import defaultdict

class InvestmentRecord():
    """Holds data for a single investment event, such as a sell or a buy.

    """

    def __init__(self, filename: str) -> None:
        self.filename = filename
        self.populate(filename)

    def populate(self, filename: str) -> None:
        """Populate instance variables from a pdf file.

        Args:
            filename: The pdf file containing data for population.
            
        """
        # Ensure requested file has a pdf extension
        if not filename.endswith(".pdf"):
            raise ValueError("Input file must have extension .pdf")
        
        with open(filename, "rb") as f:
            pdf_reader = PyPDF2.PdfFileReader(f)

            # Ensure it's a single page
            if pdf_reader.numPages > 1:
                raise Exception("Only single-page pdfs are accepted")

            # Get all the data we want
            text = pdf_reader.getPage(0).extractText()
        
        # We can handle nabTrade Contract Notes
        # These will have a first line of text as follows
        if not text.startswith("WealthHub Securities Limited"):
            raise Exception("Only nabTrade Contract Notes are accepted")

        # Grab all the record details we're interested in
        self.trade_type = re.search(r"\n(.+) [Cc]onfirmation", text).group(1)
        self.settlement_date = re.search(r"Settlement date:\n(.+)", text).group(1)
        self.confirmation_number = re.search(r"Confirmation number:\n(.+)", text).group(1)
        self.account_number = re.search(r"Account number:\n(.+)", text).group(1)
        self.hin = re.search(r"HIN:\n(.+)", text).group(1)
        # details holds multiple named groups in one big match
        details = re.search(
            r"Consideration\n(?P<quantity>.+)\n(?P<code>.+)\n(?P<security_description>[^$]+)(?P<average_price>.+)\n(?P<consideration>.+)\n", 
            text)
        self.quantity = details.group("quantity")
        self.code = details.group("code")
        self.average_price_per_share = details.group("average_price")
        self.brokerage = re.search(r"Brokerage\n(.+)", text).group(1)

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

def initialise_for_new_code(workbook: Workbook, records: list):
    # We can work on the workbook as well as the list of records in place
    # Create the sheet
    # Add the headings
    # Sort the records by date
    # Initialise the records' quantities held
    pass

def construct_investment_record_workbook(investment_records: List[InvestmentRecord]) -> Workbook:
    """ 
    """
    workbook = Workbook()
    # Hello world:
    # worksheet = workbook.active
    # worksheet["A1"] = "Hello World!"
    # workbook.save("helloworld.xlsx")

    # We will have a sheet for each code (i.e. stock exchange ticker)
    investment_records_by_code = defaultdict(list)
    for record in investment_records:
        investment_records_by_code[record.code].append(record)
   
    for code in investment_records_by_code:
        
        initialise_for_new_code(workbook, investment_records_by_code[code])
        
        # Initialise latest_date
        for record in investment_records_by_code[code]:
            # If this record begins a new financial year
                # Add a row giving the total capital gains for the past year
                # Cannot update dates in place, so return latest_date
            # Insert the basic data
            # Store the row for this record
            # If it's a sell
                # In a decoupled function:
                    # Find which buy records we are selling, and their quantities
                    # This will be FIFO, but allows other methods in future
                # Decrease the records' quantities held
                # Calculate the capital gains and insert it
                # For each record from chich some securities were sold:
                    # Fill in the date sold field in the spreadsheet



    # Clean up:
    # Delete the default, empty sheet if it exists

def display_help():
    """Print a help message for this script to the terminal.
    """
    raise Exception("To be completed.")

if __name__ == "__main__":
    if(len(sys.argv) > 1):
            if(sys.argv[1] in ["-h", "help"]):
                display_help()
                exit()
            else:
                path_to_search = sys.argv[1]
    else:
        path_to_search = "."

    investment_records = [InvestmentRecord(filename) for filename in get_contract_note_filenames(path_to_search)]

    construct_investment_record_workbook(investment_records)
    