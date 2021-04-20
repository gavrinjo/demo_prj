import os
import re
from datetime import datetime
from pathlib import Path


def dir_list(base_dir, lookup=None, extension=None, typ=None, exclude=None):

    dd = {
        "files": [],
        "directories": []
    }

    if lookup is None:
        lookup = "*"
    if extension is None:
        extension = "*"

    for obj in Path(base_dir).rglob(f"{lookup}.{extension}"):
        if exclude is not None:
            check_excluded = any(item in obj.parts for item in exclude)
            if check_excluded is not True:
                if obj.is_file():
                    dd["files"].append(obj)
                else:
                    dd["directories"].append(obj)
        else:
            if obj.is_file():
                dd["files"].append(obj)
            else:
                dd["directories"].append(obj)
    if typ == "f" or None:
        return dd["files"]
    elif typ == "d":
        return dd["directories"]


test_dir = Path("D:\\00_PRJS\\ITER\\08_Ax_Tender_Documentation")
excluded = ["Bare"]
for i in dir_list(test_dir, lookup="*_*_*_*_*_v*.*", extension="*", typ="f", exclude=excluded):
    print(i)
