import shutil
import os
from pathlib import Path

root_path = os.path.normpath("D:/00_herne/ALL_DOCS_download")
exclude = []

def path_components(path):
    folders = []
    while 1:
        path, folder = os.path.split(path)
        if folder != "":
            folders.append(folder)
        else:
            if path != "":
                folders.append(path)
            break
    folders.reverse()
    return folders


def get_list(path):
    list_path = []
    for path, dirs, files in os.walk(path):
        for filename in files:
            if "pdf" in filename:
                list_path.append(os.path.join(path, filename))
            else:
                continue
    return list_path


def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

for path, dirs, files in os.walk(root_path):
    dirs[:] = [d for d in dirs if d not in exclude]
    for filename in files:
        if "pdf" in filename:
            s = os.path.join(path, filename)
            d = os.path.normpath("\\".join(path_components(os.path.join(path, filename))[:4]))
            try:
                if os.path.isdir(s):
                    shutil.copytree(s, d)
                else:
                    shutil.copy2(s, d)
            except shutil.SameFileError as er:
                print(er)
            if Path(os.path.normpath("\\".join(path_components(os.path.join(path, filename))[4:]))).exists():
                shutil.rmtree(os.path.normpath("\\".join(path_components(os.path.join(path, filename))[4:])))
            #os.remove(os.path.normpath("\\".join(path_components(os.path.join(path, filename))[4:])))
