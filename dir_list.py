import os
import re
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
import itertools
import dir_list_r01 as dl


time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
root_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\40_Piping_Iso")
exclude = ["00_Archive", "01_Archive", "00_Document_templates"]
include_dir = ["00_Archive", "00_archive", "01_Archive", "01_archive"]
src_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project")     # source main path (needs to be join with subfolder)
dst_path = Path("J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\20_Sx_Working\\50_Workfiles")     # destination main path
hangers_path = "50_H&S_drawings"
valves_path = "90_Valves"

br_list = dl.dir_list(root_path, src_for="60*BR*", ext="pdf", exclude=exclude)
hs_list = dl.dir_list(Path.joinpath(src_path, hangers_path), "60*BQ*", "pdf", exclude=exclude)
vv_list = dl.dir_list(Path.joinpath(src_path, valves_path))     # list of valves

for valves in vv_list:
    print(valves.name)


