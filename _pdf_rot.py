from PyPDF2 import PdfFileReader

pdf_path = "J:\\32_IZ224_SIEMENS_Herne\\60_Construction\\10_Sx_Input\\30_Sx_Project_Documentation\\10_Mechanical_Engineering_Project\\50_H&S_drawings\\EGD\\Detail Pipe Support Drawing 60EGD10BQ013 Function- GR Prefab- UEY00Y\\60EGD10BQ013_516791836_Rev-.pdf"


def pdf_rotate(src, pge):
    pdf_reader = PdfFileReader(src)
    deg = pdf_reader.getPage(pge).get('/Rotate')
    page = pdf_reader.getPage(pge).mediaBox
    pt = 0.352777778
    w = round(float(page.getUpperRight_x() - page.getUpperLeft_x()) * pt)
    h = round(float(page.getUpperRight_y() - page.getLowerRight_y()) * pt)
    if w > h:
        if deg in [0, 180, None]:
            print('Landscape')
        else:
            print('Portrait')
    else:
        if deg in [0, 180, None]:
            print('Portrait')
        else:
            print('Landscape')
