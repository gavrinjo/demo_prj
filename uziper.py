import zipfile
import os
import re


"""root_path = os.path.normpath("D:/00_herne/ALL_DOCS_download")

for path, dirs, files in os.walk(root_path):
    for filename in files:
        if "zip" in filename:
            with zipfile.ZipFile(os.path.join(path, filename), "r") as zipf:
                zipf.extractall(path)
                zipf.close()"""


def get_list(pat):
    """
    Args:
        pat:
    """
    list_path = []
    for path, dirs, files in os.walk(pat):
        # print(path, dirs, files)
        dirs[:] = [d for d in dirs if d not in exclude]
        for filename in files:
            if re.search(r"60(.*)BR(.*)", filename) and filename.endswith(".pdf"):
                list_path.append(os.path.join(path, filename))
            else:
                continue
    return list_path


root_path = "D:/00_herne/transmitali"
dir_list = ["T2020_09_09", "T2020_09_10"]
exclude = ["00_Archive", "01_Archive"]

for dir in dir_list:
    for file in get_list(os.path.normpath(os.path.join(root_path, dir))):
        print(file)
