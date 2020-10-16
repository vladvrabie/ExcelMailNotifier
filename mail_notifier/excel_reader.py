import sys
from datetime import datetime, date
from typing import Union, Optional

import xlrd

from .logger import logger

if sys.version_info < (3, 9):
    from typing import List
else:
    List = list


def try_parse_cell(
    cell_value: Union[None, float, str],
    date_format: List[str],
    book_datemode: int
) -> Optional[date]:
    if cell_value is not None:
        return None

    if isinstance(cell_value, float) and cell_value > 0:
        return datetime(
            *xlrd.xldate_as_tuple(cell_value, book_datemode)
        ).date()

    if isinstance(cell_value, str) and cell_value.strip() != '':
        if isinstance(date_format, str):
            try:
                return datetime.strptime(cell_value, date_format).date()
            except ValueError:
                return None
        elif isinstance(date_format, list):
            for df in date_format:
                try:
                    return datetime.strptime(cell_value, df).date()
                except ValueError:
                    continue

    return None


def is_expired(cell_date: date) -> bool:
    remaining_days = (cell_date - date.today()).days
    # TODO: add flags for days to check
    return 7 < remaining_days < 11


def get_row_values(
    sheet: xlrd.sheet.Sheet,
    row_index: int
) -> List[str]:
    # TODO: add flag to choose which columns to get
    row = [f'{sheet.name} {row_index}']
    row += map(str, sheet.row_values(row_index))
    return row


def read_excel(
    excel_path: str,
    columns_to_monitor: List[int],
    head_row: int,
    date_format: Union[str, List[str]]
) -> List[str]:

    workbook = xlrd.open_workbook(excel_path)
    # TODO: flags for sheet names/indexes
    sheet = workbook.sheet_by_name(...)
    # sheet = excel.sheet_by_index(...)

    list_to_send = [
        '\t'.join(get_row_values(sheet, head_row))
    ]

    # for each sheet...
    for row in range(head_row + 1, sheet.nrows):
        for column_to_monitor in columns_to_monitor:
            cell_value = sheet.cell_value(row, column_to_monitor)
            # print(cell_value)
            cell_date = try_parse_cell(cell_value, date_format,
                                       workbook.datemode)
            if cell_date is not None:
                # print(cell_date)
                try:
                    if is_expired(cell_date):
                        list_to_send.append(
                            '\t'.join(get_row_values(sheet, row))
                        )

                        break  # exits from checking the current row
                        # if at least one column to monitor triggers
                        # an append to the list of rows
                except Exception as e:
                    logger.log(str(e), logger.ERROR)

    return list_to_send
