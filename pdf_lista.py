import os


def get_list(path):
    for path, dirs, files in os.walk(path):
        for filename in files:
            if "pdf" in filename:
                list_path.append(os.path.join(path, filename))
            else:
                continue
    return list_path

root_path = "D:\\01_test\\40_Piping_Iso\\NDA"
list_path = []


print(get_list(root_path))