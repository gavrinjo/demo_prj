from dir_list_r01 import dir_list
from pathlib import Path
from pdf_parser import parse, re_split, parse2, get_pdf_content_lines, parse3, parse4, parse5
import os

main_dir = Path("D:/_test_ground/_zeran")

ref_ls = [a.stem[:-3] for a in dir_list(main_dir, extension="pdf")]
con_ls = [a.stem[:-3] for a in dir_list(main_dir, extension="pdf")]

# latest = max(ref_ls, key=os.path.getctime)

test = all(map(lambda x, y: x == y, ref_ls, con_ls))

# print(test)


print(any(x in ref_ls for x in ref_ls))

file = Path(r"D:\_test_ground\_zeran\01_LBA_01_LB-HP\2018-07-25\Z214LBA25BR010_00.pdf")
# file = Path(r"D:\_test_ground\_zeran\01_LBA_01_LB-HP\2018-07-25\Z214LBA10BR010_00.pdf")
delimiters = " ", "\n"

# print(parse(file))

# print(list(filter(None, re_split(delimiters, parse(file),  maxsplit=0))))
# print(re_split(delimiters, parse(file),  maxsplit=0))
# print(parse2(file))

# print(searchInPDF(file))

"""from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextLineHorizontal

for page_layout in extract_pages(file):
    for element in page_layout:
        if isinstance(element, LTTextLineHorizontal):
            print(element.get_text())"""


# print(parse3(file))
parse5(file)

