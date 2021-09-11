"""

    pfin.py
    ~~~~~~
    z33k's personal finances

    @author: z33k

"""
from typing import Dict, List, Tuple, Union

import gspread
from enum import Enum


CREDS_FILE = "pfin_service_account.json"


class ValueRenderOption(Enum):
    FORMATTED_VALUE = 0
    UNFORMATTED_VALUE = 1
    FORMULA = 2


class ValueInputOption(Enum):
    INPUT_VALUE_OPTION_UNSPECIFIED = 0
    RAW = 1
    USER_ENTERED = 2


DataRow = List[int, float, str]
DataRows = List[DataRow]


class InputWorksheet:
    """Wrapper for input gspread's Worksheet object for easy retrieval of the needed data.
    """
    def __init__(self, worksheet: gspread.models.Worksheet) -> None:
        self._base_ws = worksheet
        self._raw_values = worksheet.get_all_values(
            value_render_option=ValueRenderOption.UNFORMATTED_VALUE.name)[4:]
        self._values = worksheet.get_all_values(
            value_render_option=ValueRenderOption.FORMATTED_VALUE)[4:]
        self._colmap = self._get_colmap()
        self._summary_col_numbers = [
            self.colmap["ordinal"],
            self.colmap["amount"],
            self.colmap["where"],
            self.colmap["who"],
            self.colmap["what_category"],
            self.colmap["what_item"],
            self.colmap["incomings"],
            self.colmap["date"],
            self.colmap["percentage"],
        ]
        self._summary_values, self._parents_summary_values = self._get_summary_values()

    def _get_colmap(self) -> Dict[str, Union[int, List[int]]]:
        """Get this worksheet's columns designations as dict.

        The output depends on number of bank account balance columns living in the input sheet.

        Example designations as of August 2021:
        =======================================
        colmap = {
            "ordinal": 1,
            "amount": 2,
            "balance": 3,
            "bank_tag": 4,
            "bank_accounts": [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
            "where": 18,
            "who": 19,
            "what_category": 20,
            "what_item": 21,
            "incomings": 22,
            "date": 23,
            "percentage": 25,
            "share": 26,
        }
        """
        row = self._raw_values[1]
        key = "gdzie"
        key_col_nr = next((i for i, v in enumerate(row, start=1) if v == key), None)
        if key_col_nr is None:
            raise ValueError(f"Invalid worksheet. No '{key}' string in the second row.")
        colmap = {
            "ordinal": 1,
            "amount": 2,
            "balance": 3,
            "bank_tag": 4,
            "bank_accounts": [*range(5, key_col_nr)],
            "where": key_col_nr,
            "who": key_col_nr + 1,
            "what_category": key_col_nr + 2,
            "what_item": key_col_nr + 3,
            "incomings": key_col_nr + 4,
            "date": key_col_nr + 5,
            "percentage": key_col_nr + 7,
            "share": key_col_nr + 8,  # this column is added to input after calculation
        }
        return colmap

    def _get_summary_values(self) -> Tuple[DataRows, DataRows]:
        # filter out data rows irrelevant for summary
        summary_values = [[v for i, v in enumerate(row, start=1)
                           if v[self.colmap["ordinal"] - 1] and v[self.colmap["amount"] - 1]
                           and v[self.colmap["what_category"] - 1] != "transfer"
                           and "rozliczenie" not in v[self.colmap["incomings"] - 1]
                           and v[self.colmap["percentage"] - 1]] for row in self._raw_values]

        # filter out data columns irrelevant for summary
        summary_values = [[v for i, v in enumerate(row, start=1)
                           if i in self.summary_col_numbers] for row in summary_values]

        # replace raw value date with formatted value date
        new_summary_values = []
        for row in summary_values:
            new_row = []
            for i, value in enumerate(row, start=1):
                if i == self.colmap["date"]:
                    value = self._values[self.colmap["date"] - 1]
                new_row.append(value)
            new_summary_values.append(new_row)

        summary_values = new_summary_values

        # calculate and add 'share' data column
        summary_values = [[v for v in row] +
                          [row[self.colmap["amount"] - 1] * row[self.colmap["percentage"]] / 100]
                          for row in summary_values]

        # divide data into summary_values and parents_summary_values
        summary_values = [row for row in summary_values
                          if not row[self.colmap["ordinal"] - 1].startswith("R")]
        parents_summary_values = [row for row in summary_values
                                  if row[self.colmap["ordinal"] - 1].startswith("R")]
        return summary_values, parents_summary_values

    @property
    def colmap(self) -> Dict[str, Union[int, List[int]]]:
        return self._colmap

    @property
    def summary_col_numbers(self) -> List[int]:
        return self._summary_col_numbers

    @property
    def summary_values(self) -> DataRows:
        return self._summary_values

    @property
    def parents_summary_values(self) -> DataRows:
        return self._parents_summary_values


def input_data(spreadsheet_name: str,
               worksheet_name: str) -> Tuple[DataRows, DataRows]:
    gc = gspread.service_account(filename=CREDS_FILE)
    ss = gc.open(spreadsheet_name)
    ws = ss.worksheet(worksheet_name)
    ws = InputWorksheet(ws)
    return ws.summary_values, ws.parents_summary_values
