import PyPDF2
import re
from re import Match
import os
from typing import List
import sys
from openpyxl import Workbook
# import defusedxml
from collections import defaultdict
from datetime import datetime
import os.path
from operator import attrgetter
from fiscalyear import FiscalDateTime

# When working with the spreadsheet, we will use these colum names
_COLUMN_NAMES = {}
i = 1
for name in [
    "Date",
    "Code",
    "Quantity",
    "Average price",
    "Transaction type",
    "Brokerage",
    "Capital gain",
    "Filename",
    "History"
]:
    _COLUMN_NAMES[name] = i
    i += 1
del i


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

        self.trade_type = re.search(r"\n(.+) [Cc]onfirmation", text).group(1)
        # We want dates as datetime for later use
        # mFund transactions omit Trade date, so use a suitable substitute
        trade_date = re.search(r"Trade date:\n(.+)", text)
        if(trade_date):
            self.trade_date = trade_date.group(1)
        else:
            self.trade_date = re.search(r"As at date:\n(.+)", text).group(1)
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

def initialise_for_new_code(workbook: Workbook, code: str, records: list):
    # We can work on the workbook as well as the list of records in place
    sheet = workbook.create_sheet(code)
    for name, column_ref in _COLUMN_NAMES.items(): 
        sheet.cell(1, column_ref).value = name

    records.sort(key=attrgetter("trade_date", "trade_type"))
    for record in records:
        if(record.trade_type.lower() in "buy", "application"):
            record.available_quantity = record.quantity
    pass

def financial_year_check(workbook: Workbook, prev_date: datetime, new_date: datetime, summaries: list) -> datetime:
    # Fiscal year is represented by the year of its end date
    '''
    if(FiscalDateTime(prev_date.year, prev_date.month, prev_date.day).fiscal_year < 
    FiscalDateTime(new_date.year, new_date.month, new_date.day).fiscal_year):
        # Add a row giving the total capital gains for the past year
        # Add summary data
        pass
    '''
    pass

def add_new_record_row(workbook: Workbook, record: InvestmentRecord):
    # Create new row and add the basic data
    # Return the row reference
    return ""

def find_records_to_sell_fifo(workbook: Workbook, quantity_to_sell: int) -> list:
    # Find which buy records we are selling, and their quantities
    # This will be FIFO, but allows other methods in future
    # Return a list of tuples containing each investment record from which we should sell, and the quantity sold
    return []

def add_sale_data(workbook: Workbook, sale_record: InvestmentRecord, recs_and_quants_to_sell: list):
    for rec,quant in recs_and_quants_to_sell:
        # Decrease the records' quantities held
        # Calculate the capital gains and insert it
        # For each record from which some securities were sold:
            # Fill in the date sold field in the spreadsheet
        pass

def add_summary_sheet(workbook: Workbook, all_fin_year_summaries: list):
    # Expect a list of (string,[(int, string, string)])
    # This represents (code, [(year, ref_to_transaction_fees, ref_to_capital_gains)])
    # e.g. [("FAIR", [(2019, "=FAIR.ASX!B4", "=FAIR.ASX!D4"), (2020, "=FAIR.ASX!B13", "=FAIR.ASX!D13")]]),
    #       ("FIL31", [(2019, "=FIL31.ASX!B9", "=FIL31.ASX!D9"), (2020, "=FIL31.ASX!B12", "=FIL31.ASX!D12")]])]
    pass

def construct_investment_record_workbook(investment_records: List[InvestmentRecord]) -> Workbook:
    workbook = Workbook()

    investment_records_by_code = defaultdict(list)
    for record in investment_records:
        investment_records_by_code[record.code].append(record)
    all_fin_year_summaries = []

    # We will have a sheet for each code
    for code in investment_records_by_code:
        records = investment_records_by_code[code]
        initialise_for_new_code(workbook, code, records)
        current_date = datetime.strptime(records[0].trade_date,"%d/%m/%Y")
        code_fin_year_summaries = []

        for record in investment_records_by_code[code]:
            financial_year_check(workbook, current_date, record.trade_date, code_fin_year_summaries)
            current_date = record.trade_date
            record.row_reference = add_new_record_row(workbook, record)
            if(record.trade_type.lower() in "sell", "redemption"):
                recs_and_quants_to_sell = find_records_to_sell_fifo(workbook, record.quantity)
                add_sale_data(workbook, record, recs_and_quants_to_sell)

        all_fin_year_summaries.append((code, code_fin_year_summaries))
    
    add_summary_sheet(workbook, all_fin_year_summaries)
    return workbook

def display_help():
    """Print a help message for this script to the terminal.
    """
    raise Exception("To be completed.")

def save_workbook(workbook: Workbook, filename: str):
    if os.path.isfile(filename):
        os.rename(filename, filename+datetime.now().strftime(".%Y-%m-%d.%H.%M.%S.bak"))    
    workbook.save(filename)    
 
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
    workbook = construct_investment_record_workbook(investment_records)
    save_workbook(workbook, "Investment_Record_Tally.xlsx")