
import shutil as sh
from datetime import datetime
from openpyxl import Workbook, load_workbook, utils
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from pathlib import Path

import logging_error as log

log_file = "error_logfile"
log = log.get_logger_st()

time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")       # date/time in format as (Y-m-d_HMS)
xl_file = Path("D:\\00_HERNE\\test\\test_copy.xlsx")     # xlsx file


def copy_xl(src, col1, col2, col3):

    wb = load_workbook(src, read_only=False)  # workbook
    ws = wb["Sheet1"]  # workbook sheet activate

    src_column = column_index_from_string(col1)
    dst_column = column_index_from_string(col2)
    src_list = next(i for i in ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=src_column, max_col=src_column))
    col_offset = dst_column - src_column

    for item in src_list:
        try:
            if not Path(item.offset(0, col_offset).value).parent.exists():
                Path(item.offset(0, col_offset).value).parent.mkdir(parents=True, exist_ok=True)
            sh.copy2(item.value, item.offset(0, col_offset).value)
            # print(str(item.value), item.offset(0, col_offset).value)
        except Exception as error:
            # log.exception(f"{item.value} --> {error}", exc_info=True)
            ws.cell(coordinate_from_string(item.coordinate)[1], column_index_from_string(col3), log.exception(f"{item.value} --> {error}", exc_info=False))
        else:
            ws.cell(coordinate_from_string(item.coordinate)[1], column_index_from_string(col3), "success")

    wb.save(src)
    wb.close()


copy_xl(xl_file, "A", "C", "E")
