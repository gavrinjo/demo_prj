import PyPDF2
import re
from io import StringIO
from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.layout import LAParams, LTTextLineHorizontal, LTTextBox
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, resolve1
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.high_level import extract_pages


def parse(file):
    """
    Args:
        file:
    """
    output_string = StringIO()
    with open(file, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams(detect_vertical=True))
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    pages = str(resolve1(doc.catalog["Pages"])["Count"])
    content = output_string.getvalue()
    return f"{pages} {content}"


def re_split(delimiter, string, maxsplit=0):
    rePattern = "|".join(map(re.escape, delimiter))
    return re.split(rePattern, string, maxsplit)


def parse2(file):
    pdfContent = []
    with open(file, "rb") as pdf_file:
        pdfReader = PyPDF2.PdfFileReader(pdf_file)
        for page in range(pdfReader.numPages):
            pdfContent.append(pdfReader.getPage(page).extractText())

    return pdfContent


def get_pdf_content_lines(pdf_file_path):
    with open(pdf_file_path, "rb") as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        for page in pdf_reader.pages:
            for text in page.extractText().splitlines():
                yield text


def parse3(file):

    output_string = StringIO()
    with open(file, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = PDFPageAggregator(rsrcmgr, laparams=LAParams(detect_vertical=True))
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        layout = []
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
            layout.append(device.get_result())

        return layout


def parse4(file):
    with open(file, "rb") as pdfFileObj:
        pdfRead = PyPDF2.PdfFileReader(pdfFileObj)
        pageObj = pdfRead.getPage(1)
        pages_text = pageObj.extractText()
        for line in StringIO(pages_text):
            print(line)


def parse5(file):

    #output_string = StringIO()
    with open(file, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = PDFPageAggregator(rsrcmgr, laparams=LAParams(detect_vertical=True, line_margin=2))
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
            layout = device.get_result()
            for x in layout:
                if isinstance(x, LTTextLineHorizontal):
                    print(x.get_text().strip())

    #pages = str(resolve1(doc.catalog["Pages"])["Count"])
    #content = output_string.getvalue()
    #return f"{pages} {content}"


