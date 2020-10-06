import PyPDF2
import re
from re import Match
import os
from typing import List
import sys
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import cell
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
        
        self.filename = filename

        # We can handle nabTrade Contract Notes
        # These will have a first line of text as follows
        if not text.startswith("WealthHub Securities Limited"):
            raise Exception("Only nabTrade Contract Notes are accepted")

        self.trade_type = re.search(r"\n(.+) [Cc]onfirmation", text).group(1)
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

        # Data type conversions
        self.trade_date = datetime.strptime(self.trade_date,"%d/%m/%Y")
        self.quantity = float(self.quantity.replace(",", ""))
        self.average_price_per_share = float(self.average_price_per_share.replace("$",""))
        self.brokerage = float(self.brokerage.replace("$",""))



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

def initialise_for_new_code(workbook: Workbook, code: str, records: list) -> Worksheet:
    # We can work on the workbook as well as the list of records in place
    sheet = workbook.create_sheet(code)
    for name, column_idx in _COLUMN_NAMES.items(): 
        sheet.cell(1, column_idx).value = name

    records.sort(key=attrgetter("trade_date", "trade_type"))
    for record in records:
        if(record.trade_type.lower() in "mfund buy", "mfund application"):
            record.available_quantity = record.quantity

    return sheet

def financial_year_check(sheet: Worksheet, row_idx: int, prev_date: datetime, new_date: datetime, summaries: list) -> int:
    # Fiscal year is represented by the year of its end date
    '''
    if(FiscalDateTime(prev_date.year, prev_date.month, prev_date.day).fiscal_year < 
    FiscalDateTime(new_date.year, new_date.month, new_date.day).fiscal_year):
        # Add a row giving the total capital gains for the past year
        # Add summary data
        pass
    '''
    return row_idx

def add_new_record_row(sheet: Worksheet, row_idx: int, record: InvestmentRecord) -> int:
    # Spreadsheet row is stored in the record object for later reference
    record.row_idx = row_idx
    sheet.cell(row_idx, _COLUMN_NAMES["Date"]).value = record.trade_date
    sheet.cell(row_idx, _COLUMN_NAMES["Code"]).value = record.code
    sheet.cell(row_idx, _COLUMN_NAMES["Quantity"]).value = record.quantity
    sheet.cell(row_idx, _COLUMN_NAMES["Average price"]).value = record.average_price_per_share
    sheet.cell(row_idx, _COLUMN_NAMES["Transaction type"]).value = record.trade_type
    sheet.cell(row_idx, _COLUMN_NAMES["Brokerage"]).value = record.brokerage
    sheet.cell(row_idx, _COLUMN_NAMES["Filename"]).value = record.filename
    return row_idx + 1

def find_records_to_sell_fifo(records: list, sale_record: InvestmentRecord) -> list:
    # We assume records is already sorted
    quantity_to_sell = sale_record.quantity
    recs_and_quants_sold = []
    for record in records:
        if record is sale_record:
            raise Exception("Insufficient buy records found for sale record " + sale_record.filename)
        if record.trade_type.lower() in ("sell","mfund redemption"):
            continue
        if record.available_quantity >= quantity_to_sell:
            recs_and_quants_sold.append((record, quantity_to_sell))
            record.available_quantity -= quantity_to_sell
            quantity_to_sell = 0
            break
        else:
            recs_and_quants_sold.append((record, record.available_quantity))
            quantity_to_sell -= record.available_quantity
            record.available_quantity = 0
            break
    return recs_and_quants_sold

def add_sale_data(sheet: Worksheet, sale_record: InvestmentRecord, recs_and_quants_to_sell: list):
    capital_gains_formula = "=(" + str(sale_record.quantity) + "*" \
    + cell.get_column_letter(_COLUMN_NAMES["Average price"]) \
    + str(sale_record.row_idx) + ")"

    for rec,quant in recs_and_quants_to_sell:
        capital_gains_formula += "-(" + str(quant) + "*" \
        + cell.get_column_letter(_COLUMN_NAMES["Average price"]) \
        + str(rec.row_idx) + ")"

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
    row_idx = 1
    
    # We will have a sheet for each code
    for code in investment_records_by_code:
        records = investment_records_by_code[code]
        sheet = initialise_for_new_code(workbook, code, records)
        current_date = records[0].trade_date
        code_fin_year_summaries = []
        # Row 1 used for heading
        row_idx = 2

        for record in investment_records_by_code[code]:
            row_idx = financial_year_check(sheet, row_idx, current_date, record.trade_date, code_fin_year_summaries)
            current_date = record.trade_date
            row_idx = add_new_record_row(sheet, row_idx, record)
            if record.trade_type.lower() in ("sell", "mfund redemption"):
                recs_and_quants_sold = find_records_to_sell_fifo(records, record)
                add_sale_data(sheet, record, recs_and_quants_sold)
            current_date = record.trade_date

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