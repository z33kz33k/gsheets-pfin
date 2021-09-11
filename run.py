"""

    run.py
    ~~~~~~
    Test Google Sheets

    @author: z33k

"""
import gspread
from pprint import pprint

gc = gspread.service_account(filename="pfin_service_account.json")

# Find a workbook by name and open the first sheet
# Make sure you use the right name here
sheet = gc.open("pfin_2021").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_records()
pprint(list_of_hashes)

# so far used gspread API methods:
# Spreadsheet.worksheet()
# Spreadsheet.duplicate_sheet()
# Spreadsheet.del_worksheet()
# Worksheet.batch_clear()
# Worksheet.delete_rows()
# Worksheet.update_title()
# Worksheet.get_all_values()
