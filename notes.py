import sys
import PyPDF2, traceback
import pprint
from subprocess import call

try:
    src = sys.argv[1]
except:
    src = r'D:/00_herne/50_Workfiles/EKT/60EKT20BR011/60EKT20BR011_517641871_RevB.pdf'

# put the role into the rst file
print('.. role:: slide-title')
print('')

input1 = PyPDF2.PdfFileReader(open(src, "rb"))
nPages = input1.getNumPages()

for i in range(nPages):
    # get the data from this PDF page (first line of text, plus annotations)
    page = input1.getPage(i)
    text = page.extractText()

    print(':slide-title:`' + text.splitlines()[0] + '`')
    print('')

    try:
        for annot in page:
            # Other subtypes, such as /Link, cause errors
            subtype = annot.getObject()
            print(annot.getObject())
            print('')
    except:
        # there are no annotations on this page
        pass

    print('')
