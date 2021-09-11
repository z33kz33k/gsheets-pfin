"""

    pfin.py
    ~~~~~~
    z33k's personal finances

    @author: z33k

"""
from typing import List, Tuple, Union

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


SummaryValuesList = List[List[Union[int, float, str]]]


class InputWorksheet:
    """Wrapper for input gspread's Worksheet object for easy retrieval of the needed data.
    """
    AMOUNT_COL_NR = 2

    def __init__(self, worksheet: gspread.models.Worksheet) -> None:
        self._base_ws = worksheet
        self._raw_values = worksheet.get_all_values(
            value_render_option=ValueRenderOption.UNFORMATTED_VALUE.name)[4:]
        self.summary_col_numbers, self.category_col_nr, self.incomings_col_nr, \
            self.percentage_col_nr = self._get_worksheet_params()
        self.summary_values, self.parents_summary_values = self._get_summary_values()

    def _get_worksheet_params(self) -> Tuple[List[int], int, int, int]:
        """Get this worksheet parameters.

        The output depends on number of bank account balance columns living in the input sheet.

        Example parameters for August 2021:

        summary_col_numbers = [1, 2, 18, 19, 20, 21, 22, 23, 24, 25]
        category_col_nr = 20
        percentage_col_nr = 25
        """
        row = self._raw_values[1]
        key = "gdzie"
        key_col_nr = next((i for i, v in enumerate(row, start=1) if v == key), None)
        if key_col_nr is None:
            raise ValueError(f"Invalid worksheet. No '{key}' string in the second row.")
        rng = [*range(key_col_nr, key_col_nr + 8)]
        return [1, 2] + rng, rng[2], rng[4], rng[-1]

    def _get_summary_values(self) -> Tuple[SummaryValuesList, SummaryValuesList]:
        summary_values = [[v for i, v in enumerate(row, start=1)
                           if v[self.AMOUNT_COL_NR - 1]
                           and v[self.category_col_nr - 1] != "transfer"
                           and "rozliczenie" not in v[self.incomings_col_nr - 1]
                           and v[self.percentage_col_nr - 1]] for row in self._raw_values]

        summary_values = [[v for i, v in enumerate(row, start=1)
                           if i in self.summary_col_numbers] for row in summary_values
                          if row[0]]

        summary_values = [row for row in summary_values if not row[0].startswith("R")]
        parents_summary_values = [row for row in summary_values if row[0].startswith("R")]
        return summary_values, parents_summary_values


def input_data(spreadsheet_name: str,
               worksheet_name: str) -> Tuple[SummaryValuesList, SummaryValuesList]:
    gc = gspread.service_account(filename=CREDS_FILE)
    ss = gc.open(spreadsheet_name)
    ws = ss.worksheet(worksheet_name)
    ws = InputWorksheet(ws)
    return ws.summary_values, ws.parents_summary_values
