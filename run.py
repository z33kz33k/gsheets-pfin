"""

    run.py
    ~~~~~~
    Test Google Sheets

    @author: z33k

"""
import gspread
from pfin import InputWorksheet, CREDS_FILE, OutputWorksheet

spreadsheet_name, worksheet_name = "pfin_2021", "202108"
gc = gspread.service_account(filename=CREDS_FILE)
ss = gc.open(spreadsheet_name)
ws = ss.worksheet(worksheet_name)
ws = InputWorksheet(ws, verbose=True)
sv, psv = ws.summary_values, ws.parents_summary_values

t = ss.worksheet("template_sandbox")
ows = OutputWorksheet(t, sv, psv, verbose=True)
ows.upload_data()




