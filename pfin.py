"""

    pfin.py
    ~~~~~~
    z33k's personal finances

    @author: z33k

"""
from typing import Dict, List, Tuple, Union

import gspread
from gspread_formatting import CellFormat, Color, ColorStyle, NumberFormat, Padding, TextFormat, \
    format_cell_range, batch_updater
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


DataRow = List[Union[int, float, str]]
DataRows = List[DataRow]


class InputWorksheet:
    """Wrapper of input gspread's Worksheet object for easy retrieval of the needed data.
    """
    def __init__(self, worksheet: gspread.Worksheet, verbose=False) -> None:
        self._base_ws, self.verbose = worksheet, verbose
        self._raw_values_full = worksheet.get_all_values(
            value_render_option=ValueRenderOption.UNFORMATTED_VALUE.name)
        self._raw_values = self._raw_values_full[4:]
        self._values_full = worksheet.get_all_values(
            value_render_option=ValueRenderOption.FORMATTED_VALUE.name)
        self._values = self._values_full[4:]
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
            self.colmap["separator"],
            self.colmap["percentage"],
            self.colmap["share"],
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
            "separator": 24,
            "percentage": 25,
            "share": 26,
        }
        """
        row = self._raw_values_full[1]
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
            "separator": key_col_nr + 6,
            "percentage": key_col_nr + 7,
            "share": key_col_nr + 9,  # this column is added to input after calculation
        }
        return colmap

    def _get_summary_values(self) -> Tuple[DataRows, DataRows]:
        # replace raw value date with formatted value date
        summary_values = []
        for index, row in enumerate(self._raw_values):
            new_row = []
            for i, value in enumerate(row, start=1):
                if i == self.colmap["date"]:
                    value = self._values[index][self.colmap["date"] - 1]
                new_row.append(value)
            summary_values.append(new_row)
        if self.verbose:
            print("Replacing raw-value dates with formatted-value dates completed.")

        # filter out data rows irrelevant for summary
        summary_values = [row for row in summary_values
                          if row[self.colmap["ordinal"] - 1] and row[self.colmap["amount"] - 1]
                          and row[self.colmap["what_category"] - 1] != "transfer"
                          and "rozliczenie" not in row[self.colmap["incomings"] - 1]
                          and row[self.colmap["percentage"] - 1]]
        if self.verbose:
            print("Filtering out data rows irrevelant for the summary completed.")

        # re-calculate 'share' data column
        for row in summary_values:
            amount = row[self.colmap["amount"] - 1]
            percentage = row[self.colmap["percentage"] - 1]
            share = amount * percentage / 100
            row[self.colmap["share"] - 1] = share
        if self.verbose:
            print("Re-calculating 'share' data column completed.")

        # filter out data columns irrelevant for summary
        summary_values = [[v for i, v in enumerate(row, start=1)
                           if i in self.summary_col_numbers] for row in summary_values]
        if self.verbose:
            print("Filtering out data columns irrevelant for the summary completed.")

        # divide data into summary_values and parents_summary_values
        parents_summary_values = [row for row in summary_values
                                  if row[self.colmap["ordinal"] - 1].startswith("R")]
        summary_values = [row for row in summary_values
                          if not row[self.colmap["ordinal"] - 1].startswith("R")]
        if self.verbose:
            print("Input data retrieval completed.")
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
               worksheet_name: str, verbose=False) -> Tuple[DataRows, DataRows]:
    gc = gspread.service_account(filename=CREDS_FILE)
    ss = gc.open(spreadsheet_name)
    ws = ss.worksheet(worksheet_name)
    ws = InputWorksheet(ws, verbose=verbose)
    return ws.summary_values, ws.parents_summary_values


class OutputWorksheet:
    """Wrapper of output gspread's Worksheet object for easy data upload and manipulation.
    """
    FIRST_DATAROW = 3
    DATE_COL, PERCENT_COL, RESULT_COL = "H", "J", "K"
    FIRST_PERCENT_CELL = f"{PERCENT_COL}{FIRST_DATAROW}"
    PERCENT_FORMAT = CellFormat(
        backgroundColor=Color(1, 1, 1),
        backgroundColorStyle=ColorStyle(rgbColor=Color(1, 1, 1)),
        horizontalAlignment="RIGHT",
        hyperlinkDisplayType="PLAIN_TEXT",
        numberFormat=NumberFormat(type="NUMBER", pattern='#,##0"%"'),
        padding=Padding(2, 3, 2, 3),
        textFormat=TextFormat(
            foregroundColor=Color(red=0.0, green=0.0, blue=0.0, alpha=1.0),
            bold=False,
            fontFamily="Roboto Mono",
            fontSize=8,
            foregroundColorStyle=ColorStyle(rgbColor=Color(red=0.0, green=0.0, blue=0.0,
                                                           alpha=1.0)),
            italic=False,
            strikethrough=False,
            underline=False
        ),
        verticalAlignment="BOTTOM",
        wrapStrategy="WRAP"
    )
    DATE_FORMAT = CellFormat(
        backgroundColor=Color(1, 1, 1),
        backgroundColorStyle=ColorStyle(rgbColor=Color(1, 1, 1)),
        horizontalAlignment="RIGHT",
        hyperlinkDisplayType="PLAIN_TEXT",
        numberFormat=NumberFormat(type="DATE", pattern="yyyy-mm-dd"),
        padding=Padding(2, 3, 2, 3),
        textFormat=TextFormat(
            foregroundColor=Color(red=0.0, green=0.0, blue=0.0, alpha=1.0),
            bold=False,
            fontFamily="Roboto Mono",
            fontSize=8,
            foregroundColorStyle=ColorStyle(rgbColor=Color(red=0.0, green=0.0, blue=0.0,
                                                           alpha=1.0)),
            italic=False,
            strikethrough=False,
            underline=False
        ),
        verticalAlignment="BOTTOM",
        wrapStrategy="WRAP"
    )

    def __init__(self, worksheet: gspread.Worksheet, summary_values: DataRows,
                 parents_summary_values: DataRows, verbose=False) -> None:
        self._base_ws, self.verbose = worksheet, verbose
        self._summary_values = summary_values
        self._parents_summary_values = parents_summary_values

    @property
    def summary_values(self) -> DataRows:
        return self._summary_values

    @property
    def parents_summary_values(self) -> DataRows:
        return self._parents_summary_values

    def _upload_values(self, values: DataRows, first_row: int) -> None:
        if self.verbose:
            print("Commencing summary values upload.")
        self._base_ws.insert_rows(values, row=first_row)

        # correct Percentage and Date columns formatting
        end_range_row = first_row + len(values) - 1
        first_percent_cell = f"{self.PERCENT_COL}{first_row}"
        end_percent_cell = f"{self.PERCENT_COL}{end_range_row}"
        percent_cell_range = f"{first_percent_cell}{end_percent_cell}"  # ex. "J3:J36"
        format_cell_range(self._base_ws, percent_cell_range, self.PERCENT_FORMAT)
        format_cell_range(self._base_ws,
                          percent_cell_range.replace(self.PERCENT_COL, self.DATE_COL),
                          self.DATE_FORMAT)
        if self.verbose:
            print("Correction of Percentage and Date columns' formatting completed.")

        # delete the empty trailing row
        self._base_ws.delete_rows(end_range_row + 1)
        if self.verbose:
            print("Trailing empty row deleted.")

        # update the result's cell formula
        result_row = end_range_row + 1
        label = f"{self.RESULT_COL}{result_row}"
        result_cell = self._base_ws.acell(label)
        start_range_lbl = f"{self.RESULT_COL}{first_row}"
        end_range_lbl = f"{self.RESULT_COL}{end_range_row}"
        result_cell.value = f"=SUM({start_range_lbl}:{end_range_lbl})"
        self._base_ws.update_cells([result_cell],
                                   value_input_option=ValueInputOption.USER_ENTERED.name)
        if self.verbose:
            print("Updating the result cell's formula completed.")

    def upload_data(self) -> None:
        self._upload_values(self.summary_values, self.FIRST_DATAROW)
        first_parents_row = self.FIRST_DATAROW + len(self.summary_values) + 2
        self._upload_values(self.parents_summary_values, first_parents_row)

    def duplicate(self, sheetname: str) -> None:
        self._base_ws.duplicate(self._base_ws.id, new_sheet_name=sheetname)
