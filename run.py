"""

    run.py
    ~~~~~~
    Test Google Sheets

    @author: z33k

"""
import gspread
from pprint import pprint

gc = gspread.service_account(filename="gd_service_account.json")

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = gc.open("pfin_2021").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_records()
pprint(list_of_hashes)
