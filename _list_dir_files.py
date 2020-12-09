import re
import os
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
import dir_list_r01 as dl


time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
# root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\50_H&S_drawings")
root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\20_Sx_Working\\50_Workfiles")
exclude = ["00_Archive", "01_Archive", "00_Document_templates", "01_Deleted lines", "SKID"]

# ls = dl.dir_list(root_path, ext="pdf", exclude=exclude)     # list of valves

fn_list = list()
pattern = r"(\d\d)(BR)"
for i in dl.dir_list(root_path, obj_type="d", exclude=exclude):
    if re.search(pattern, i.name):
        fn_list.append(i)
"""
wb_save_path = Path("D:\\00_HERNE\\_tracking")
wb_save_name = "_wf_list"
wb = Workbook()     # workbook
ws = wb.active      # workbook sheet activate
r = 1       # initial row number
# patt = r"(\d{9})"

for file in fn_list:
    ws.cell(r, 1, file.parts[-3])                   # system
    ws.cell(r, 2, file.parts[-2])                   # pipe
    ws.cell(r, 3, file.stem.split("_")[0])          # support point
    # ws.cell(r, 4, file.stem.split("_")[1])          # support point unid
    # ws.cell(r, 5, file.stem.split("_")[-1][-1])     # revision
    # ws.cell(r, 6, datetime.fromtimestamp(os.path.getmtime(str(file))).strftime("%d.%m.%Y %H:%M"))
    # ws.cell(r, 7, file.name)                        # filename
    ws.cell(r, 8, str(file))                        # file path
    r += 1

wb.save(Path.joinpath(wb_save_path, f"{wb_save_name}_{time_now}.xlsx"))
wb.close()
"""