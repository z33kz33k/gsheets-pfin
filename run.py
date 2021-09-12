"""

    run.py
    ~~~~~~
    Run the script for debugging and testing.

    @author: z33k

"""
import gspread
from pfin import InputWorksheet, CREDS_FILE, OutputWorksheet

spreadsheet_name, input_worksheet_name = "pfin_2021", "202108"
gc = gspread.service_account(filename=CREDS_FILE)
ss = gc.open(spreadsheet_name)
ws = ss.worksheet(input_worksheet_name)
ws = InputWorksheet(ws, verbose=True)
sv, psv = ws.summary_values, ws.parents_summary_values

output_worksheet = ss.worksheet("template_sandbox")
ows = OutputWorksheet(output_worksheet, sv, psv, verbose=True)
ows.upload_data()




