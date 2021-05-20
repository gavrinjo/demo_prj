from dir_list_r01 import dir_list
from pathlib import Path
import os

main_dir = Path("D:/_test_ground/_zeran")

ref_ls = [a.stem[:-3] for a in dir_list(main_dir, extension="pdf")]
con_ls = [a.stem[:-3] for a in dir_list(main_dir, extension="pdf")]

# latest = max(ref_ls, key=os.path.getctime)

test = all(map(lambda x, y: x == y, ref_ls, con_ls))

# print(test)


print(any(x in ref_ls for x in ref_ls))

