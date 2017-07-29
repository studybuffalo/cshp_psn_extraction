"""
    Conversion Plans

    .html	wkhtmltopdf             Confirmed
    .doc	comtypes                Confirmed
    .txt	comtypes                Confirmed
    .pdf	N/A                     Confirmed
    .docx	comtypes                Confirmed
    .png	FPDF                    Confirmed
    .ppt	comtypes                Confirmed
    .PDF	N/A                     Confirmed
    .xps	***Manual***
    .DOC	comtypes                Confirmed
    .xls	comtypes                Confirmed
    .mht	***Manual***
    .xlsx	comtypes                Confirmed
    .DOCX	comtypes                Confirmed
    .gif	FPDF                    Confirmed
    .jpg	FPDF                    Confirmed
    .JPG	FPDF                    Confirmed
    .dot	comtypes                Confirmed
    .pptx	comtypes                Confirmed
    .pub	***Manual***
    .PNG	FPDF                    Confirmed
    .docm	comtypes                Confirmed
    .rtf	comtypes                Confirmed
    .jpeg	FPDF                    Confirmed
    .csv	comtypes                Confirmed
    .msg	***Manual***            
    .wpd	comtypes                Confirmed
    .bmp	***Manual***
    .xltx	comtypes                Confirmed
"""
import sys
import os
import comtypes.client

wdFormatPDF = 17
pptFormatPDF = 32
xlsFormatPDF = 57

in_file = os.path.abspath(sys.argv[1])
out_file = os.path.abspath(sys.argv[2])

"""
# .doc .docx .dot .docm .txt .rtf .wpd
word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
"""
"""
# .ppt .pptx
ppt = comtypes.client.CreateObject("Powerpoint.Application")
deck = ppt.Presentations.Open(in_file)
deck.SaveAs(out_file, FileFormat=pptFormatPDF)
deck.Close()
ppt.Quit()
"""
"""
# .xls .xlsx .xltx .csv
xls = comtypes.client.CreateObject("Excel.Application")
wb = xls.Workbooks.Open(in_file)
numWS = wb.Worksheets.Count

for i in range(1, numWS+1):
    try:
        wb.Worksheets[i].SaveAs("%s%s" % (out_file, i), FileFormat=xlsFormatPDF)
    except Exception as e:
        None

wb.Close()
xls.Quit()
"""

# .png
from fpdf import FPDF
from PIL import Image
image = Image.open(in_file)
height = image.size[0]
width = image.size[1]
pdf = FPDF(unit="pt", format=[height, width])
pdf.add_page()
pdf.image(in_file, 0, 0, 0, 0)
pdf.output(out_file)