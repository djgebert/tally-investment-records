import re

text = 'Class DescriptionRate per UnitParticipating UnitsGross AmountOrdinary Units23.340806 cents2,069$482.92Net Amount:$482.92Residual balance brought forward from your Plan account:$15.08Total amount available for reinvestment:$498.00This amount has been applied to 27 units at $18.093906 per unit:$488.54Residual balance carried forward in your Plan account:$9.46Number of ordinary units held prior to allotment:2,069Ordinary units allotted this distribution:27^Total holding of ordinary units after the allotment:2,0961301011002301230112111333101313110313 031  064644MR DANIEL JAMES GEBERT & MRSROSE ESTELLE GEBERT61 MARRIOTT STPARKDALE VIC 3195ASX Code: FAIRDistribution AdvicePayment date:17 January 2020Record date:3 January 2020Reference no.:X*******9436TFN/ABN RECEIVED AND RECORDED^The ªTotal holding of ordinary units after the allotmentº shown above may not be the current holding balance (it does notinclude any Ex distribution transfers registered after Ex date 2 January 2020, or any transfers registered since the Record date).Neither BetaShares nor Link Market Services Limited will be liable for any losses incurred by any person who relies on theholding shown without making their own adjustments for any transactions.ARSN 608 057 996BetaShares Australian Sustainability Leaders ETF (ASX Code: FAIR)Distribution statement for the period ended 31 December 2019A distribution payment has been made in respect of your units as at the record date. The final details of the distributioncomponents (including any non-assessable amounts) will be advised in the Attribution Managed Investment Trust MemberAnnual (AMMA) Statement for the year ending 30 June 2020.Visit our investor website at www.linkmarketservices.com.au where you can view and change your details, including electing to receivedistribution notifications by email going forward.*S0646441Q01*'

trade_date = re.search(r"Payment date:(.+)Record date:", text).group(1)

if trade_date:
    print(trade_date)

'''# Requires tesseract and popplar to be installed and added to path separately

import textract
filename = r"D:\Documents\Git\tally-investment-records\sample_data\VDGR_Reinvestment_Plan_Advice_2019_10_16.pdf"
text = textract.process(filename, method='tesseract', language='eng')
with open("Sample_VDGR_Reinvestment_Plan_Advice.txt", "wb") as f:
    f.write(text)'''


'''import PyPDF2

with open(".\sample_data\VDGR_Reinvestment_Plan_Advice_2019_10_16.pdf", "rb") as f:
    pdf_reader = PyPDF2.PdfFileReader(f)

    # Get all the data we want
    print(pdf_reader.getDocumentInfo())
    text = pdf_reader.getPage(0).extractText()
    print(text)

print("Finished.")'''