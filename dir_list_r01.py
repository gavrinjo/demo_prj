import os
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from pathlib import Path


def dir_list(src, obj_type=None, src_for=None, ext=None, exclude=None):

    if src_for is None:
        src_for = "*"
    if ext is None:
        ext = "*"

    if obj_type == "f":
        obj_type = f"/{src_for}.{ext}"
    elif obj_type is None or "d":
        obj_type = ""

    f_list = list()

    for obj in src.glob(f"**{obj_type}"):

        if exclude is None:
            f_list.append(obj)
        elif exclude is not None:
            check_excluded = any(item in obj.parts for item in exclude)
            if check_excluded is not True:
                f_list.append(obj)
            else:
                pass

    return f_list

