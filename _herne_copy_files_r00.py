
# TODO copy KKS (BR(branch), BQ(support), AA(valve)) from source input folder to WF(work file) folders

import dir_list_r01 as dl
import logging_error as log
import pdf_parser

import os
import re
import shutil as sh
from datetime import datetime
from openpyxl import Workbook, load_workbook
from pathlib import Path


log_file = "error_logfile"
log = log.get_logger(log_file)

time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")       # date/time in format as (Y-m-d_HMS)
# root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\40_Piping_Iso")     # main path
temp_path = Path("D:\\00_herne\\test\\workfiles")       # temporary path (for testing purposes)
exclude_dir = ["00_Document_templates", "00_Archive", "00_archive", "01_Archive", "01_archive"]     # excluded folders (these are skipped)
search_for = "60*BR*"   # search pattern
search_ext = "pdf"      # extension to look for


src_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project")     # source main path (needs to be join with subfolder)
dst_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\20_Sx_Working\\50_Workfiles")     # destination main path

hangers_path = "50_H&S_drawings"
valves_path = "90_Valves"


br_list = dl.dir_list(temp_path, search_for, search_ext, exclude_dir)    # list of required files
hs_list = dl.dir_list(Path.joinpath(src_path, hangers_path), "96*BQ*", "pdf", exclude_dir)      # list of hanger & support files
vv_list = dl.dir_list(Path.joinpath(src_path, valves_path))     # list of valves

for file in br_list:
    try:
        raw_pdf_data = pdf_parser.parse(file)
        raw_str_list = list(map(str, raw_pdf_data.split()))
        pattern = r"(\d\d)('BR'|'BQ')"
        for string in raw_str_list:
            if re.search(pattern, string):
                for br in br_list:
                    if br.stem.split("_")[0] == string:
                        sh.copy2(br, Path.joinpath(dst_path, br.parts[-3], br.parts[-2], file.name))
                for bq in hs_list:
                    if bq.stem.split("_")[0] == string:
                        sh.copy2(bq, Path.joinpath(dst_path, bq.parts[-3], bq.parts[-2], file.name))
        sh.copy2(file, Path.joinpath(dst_path, file.parts[-3], file.parts[-2], file.name))

