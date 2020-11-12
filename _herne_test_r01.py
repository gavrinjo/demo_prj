
# TODO izlistanje ventila it 00_Archive foldera

import dir_list_r01 as dl
import logging_error as log
import pdf_parser

import os
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from pathlib import Path

log_file = "error_logfile"
log = log.get_logger(log_file)

time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")       # date/time in format as (Y-m-d_HMS)
root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\40_Piping_Iso")     # main path
exclude_dir = ["00_Document_templates"]     # excluded folders (these are skipped)
include_dir = ["00_Archive", "00_archive", "01_Archive", "01_archive"]      # included folders (search only in these)
search_for = "60*BR*"   # search pattern
search_ext = "pdf"      # extension to look for

dlist = dl.dir_list(root_path, search_for, search_ext, exclude_dir, include_dir)    # list of required files

wb_save_path = Path("D:\\00_HERNE\\test\\")     # workbook save path
wb_file_name = "valves_from_archive_folder"     # workbook save filename
wb = Workbook()     # workbook
ws = wb.active      # workbook sheet activate
r = 1       # initial row number

for file in dlist:
    try:
        raw_pdf_data = pdf_parser.parse(file)
        raw_str_list = list(map(str, raw_pdf_data.split()))
        pattern = fr"(\d){'AA'}"
        for string in raw_str_list:
            if re.search(pattern, string):
                ws.cell(r, 1, file.stem.split("_")[0])
                ws.cell(r, 2, file.stem.split(".")[-1])
                ws.cell(r, 3, datetime.fromtimestamp(os.path.getmtime(file)).strftime("%d.%m.%Y %H:%M"))
                ws.cell(r, 4, file.name)
                ws.cell(r, 5, str(file))
                ws.cell(r, 6, string)
        r += 1
    except Exception as error:
        log.exception(f"{file} --> {error}", exc_info=True)

wb.save(Path.joinpath(wb_save_path, f"{wb_file_name}_{time_now}.xlsx"))
wb.close()
