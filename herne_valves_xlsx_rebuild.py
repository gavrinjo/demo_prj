from io import StringIO
import os
import re
from datetime import datetime
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from openpyxl import Workbook, load_workbook
import logging_error as log


def parse_pdf(file):
    output_string = StringIO()
    with open(file, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    return output_string.getvalue()


# root_path = os.path.normpath("J:/32_IZ224_SIEMENS_Herne/60_Construction/10_Sx_Input/30_Sx_Project_Documentation/10_Mechanical_Engineering_Project/40_Piping_Iso")
# dir_list = []
# subdir_list = next(os.walk(root_path))[1]
save_path = os.path.normpath("D:\\00_HERNE\\test\\")
# sys_name = os.path.basename(root_path)
file_name = "mda_valves_rebuilded"
time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
#pdf_filename = 'I:/00_PROJECTS/32_IZ224_SIEMENS_Herne/60_Construction/05_Sx_Overview/01_PLAN_contracted/2.1.3_TimeSchedule_Piping_RevB.pdf'
search_for = "BR"
exclude = ["00_Archive", "01_Archive", "00_Document_templates"]
log_file = "log_error"
# list_path = []
#output_string = StringIO()
log = log.get_logger(log_file)

xls_file = "D:\\00_HERNE\\test\\raw_valve_list.xlsx"
wb = load_workbook(xls_file, read_only=False)
sh1 = wb["Sheet1"]   # isometric line list

valve_list = next(i for i in sh1.iter_cols(min_row=1, max_row=207, min_col=1, max_col=1))    # valve kks list

wb_new = Workbook()
sh_new = wb_new.active

c = 1
r = 1

try:
    for kks in valve_list:
        for cell in next(i for i in sh1[f'D{c}':f'HH{c}']):
            if cell.col_idx == 4 and cell.value is None:
                sh_new.cell(r, 1, kks.value)
                sh_new.cell(r, 2, kks.offset(0, 1).value)
                sh_new.cell(r, 3, kks.offset(0, 2).value)
                sh_new.cell(r, 4, "fale kksovi")
            elif cell.col_idx >= 4 and cell.value is not None:
                sh_new.cell(r, 1, kks.value)
                sh_new.cell(r, 2, kks.offset(0, 1).value)
                sh_new.cell(r, 3, kks.offset(0, 2).value)
                sh_new.cell(r, 4, cell.value)
            else:
                r -= 1
            r += 1
        c += 1

except Exception as err:
    log.exception(f"{err}", exc_info=True)

wb_new.save(os.path.join(save_path, f"{file_name}_{time_now}.xlsx"))
wb_new.close()
wb.close()
