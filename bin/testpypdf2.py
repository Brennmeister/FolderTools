import PyPDF2 as pypdf
import numpy as np

def mm2units(length_mm):
    return length_mm*2.54/161.3*180

def units2mm(length_units):
    return length_units/mm2units(1)

fin = open('d:/tmp.pdf', 'rb')
fout = open('d:/tmp_out.pdf', 'wb')

pdfReader = pypdf.PdfFileReader(fin)
pdfWriter = pypdf.PdfFileWriter()

inPage = pdfReader.getPage(0)

outPage = pdfWriter.addBlankPage(width=mm2units(297), height=mm2units(210))
rows = [0,1]
cols = [0,1,2]
insert_width = units2mm(float(inPage.mediaBox[2]))
insert_heigth = units2mm(float(inPage.mediaBox[3]))

for r in rows:
    for c in cols:
    outPage.mergeTranslatedPage(inPage, tx=0.0 + mm2units(insert_width)*c, ty=0.0 + mm2units(insert_heigth)*r)



pdfWriter.write(fout)
fin.close()
fout.close()