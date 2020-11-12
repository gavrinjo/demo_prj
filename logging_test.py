import os
import re
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import datetime
from win32com import client
import logging
import logging_error

path = "D:/00_herne/01_py_script_export/logs"
file_name = "testiranje"

logger = logging_error.get_logger(file_name)

# logging.basicConfig(level=logging.DEBUG, format='%(levelname)s - %(asctime)s - %(message)s', datefmt='T(%d.%m.%Y. %H:%M)')

a = 2
b = 0
for i in range(a):
    try:
        print(i / b)
    except Exception as error:
        logger.exception(f"{error}", exc_info=False)
