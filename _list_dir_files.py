import os
import re
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
import itertools
import dir_list_r01 as dl


time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\90_Valves")
exclude = ["00_Archive", "01_Archive", "00_Document_templates"]

ls = dl.dir_list(root_path, ext="pdf", exclude=exclude)     # list of valves

wb_save_path = Path("D:\\00_HERNE\\test")
wb_save_name = "_valves_list"
wb = Workbook()     # workbook
ws = wb.active      # workbook sheet activate
r = 1       # initial row number

for v in ls:
    ws.cell(r, 1, v.parts[-2])
    ws.cell(r, 2, v.name)
    ws.cell(r, 3, str(v))
    r += 1

wb.save(Path.joinpath(wb_save_path, f"{wb_save_name}_{time_now}.xlsx"))
wb.close()

