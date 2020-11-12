from bs4 import BeautifulSoup as bs
from requests import get
from configparser import ConfigParser as cp
import xlrd
from contextlib import closing
import json
import re
from argparse import ArgumentParser as ap
import xml.etree.ElementTree as et

parser = ap(prog="n_scraper", description="nekakav opis, onako nek se nadje")
parser.add_argument('-z', help='ime zupanije!')
parser.add_argument('-c', help='ime grada!')
parser.parse_args()
args = parser.parse_args()

tree = et.parse('test.xml')
root = tree.getroot()


def city():
    for county in root:
        if county.attrib["name"] == args.z:
            return county.attrib["code"]
        for child in county:
            if child.attrib["name"] == args.c:
                return child.attrib["code"]


print(city())

def read_ini(src):
    conf = cp()
    conf.read(src, encoding='utf8')
    for i in conf['county']:
        print(i)


def get_url_base(src):
    # wb = xlrd.open_workbook(src)
    # sheet = wb.sheet_by_index(0)
    # sheet.cell_value(0, 0)

    grad = 1153
    pmin = 5000
    pmax = 13000
    includeOtherCategories = 1
    livingArea_min = 70
    livingArea_max = 130
    adsWithImages = 1
    balconyInfo = "allthree"
    buildingInfo_lift = 1

    """
    grad = sheet.cell_value(4, 4)
    pmin = int(sheet.cell_value(5, 4))
    pmax = int(sheet.cell_value(5, 5))
    includeOtherCategories = sheet.cell_value(6, 4)
    livingArea_min = int(sheet.cell_value(7, 4))
    livingArea_max = int(sheet.cell_value(7, 5))
    adsWithImages = sheet.cell_value(8, 4)
    balconyInfo = sheet.cell_value(9, 4)
    buildingInfo_lift = sheet.cell_value(10, 4)
    """
    return f"https://www.njuskalo.hr/prodaja-stanova/{grad}?price%5Bmin%5D={pmin}&price%5Bmax%5D={pmax}"


# print(get_url("nj_lookup.xlsx"))


"""def url_list(src):
    list = get(src)"""

#read_ini('njuskalo.ini')
