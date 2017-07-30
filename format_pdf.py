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

    Process Planning
    1) Convert all relevant file types to pdf
        Number them in order of merging
            Thread
            Attachments in order of appearance
        Attachments should be accessible in attachments folder
        Will need to go in manually for specified file types and update 
         attachments prior to running conversion and merge for other items
    2) Merge files together
"""
class Attachment:
    def __init__(self, title, file):
        self.title = title
        self.file = file

def get_attachment_names(thread):
    aTxt = thread.child("attachments.txt")
    attachments = []

    if aTxt.exists():
        with open(aTxt, "r") as list:
            for a in list:
                split = a.split("    |    ")
                attachments.append(Attachment(split[1], split[2]))

    # Remove the attachments file
    aTxt.remove()

    return attachments


def convert_word(file):
    # Converts: doc docx dot docm txt rtf wpd
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()
    word.Quit()

def convert_ppt(file):
    # Converts: ppt pptx
    ppt = comtypes.client.CreateObject("Powerpoint.Application")
    deck = ppt.Presentations.Open(in_file)
    deck.SaveAs(out_file, FileFormat=32)
    deck.Close()
    ppt.Quit()

def convert_xls(file):
    # Converts: xls xlsx xltx csv
    # Note: Saves each worksheet as separate file
    xls = comtypes.client.CreateObject("Excel.Application")
    wb = xls.Workbooks.Open(in_file)
    numWS = wb.Worksheets.Count

    for i in range(1, numWS+1):
        try:
            wb.Worksheets[i].SaveAs("%s%s" % (out_file, i), FileFormat=57)
        except Exception as e:
            None

    wb.Close()
    xls.Quit()

def convert_image(file):
    # Converts: png gif jpg jpeg
    image = Image.open(in_file)
    height = image.size[0]
    width = image.size[1]
    pdf = FPDF(unit="pt", format=[height, width])
    pdf.add_page()
    pdf.image(in_file, 0, 0, 0, 0)
    pdf.output(out_file)

import sys
from unipath import Path, DIRS, FILES
import comtypes.client
from PyPDF2 import PDFFileMerger
from fpdf import FPDF
from PIL import Image

# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

# Cycle through the PSN directories (main eForum directories)
fPSN = root.child("formattedPSN")
pPSN = root.child("pdfPSN")

for fForum in fPSN.listdir(filter=DIRS):
    print ("\nAccessing %s Forum" % fForum.components()[-1])
    print ("----------------------------------------------------------------------")

    # Create the new forum folder
    pForum = pPSN.child(fForum.components()[-1])
    pForum.mkdir()

    # Create a temp folder to hold files
    temp = pForum.child("temp")
    temp.mkdir()

    # Cycle through each thread
    for fThread in forum.listdir(filter=DIRS):
        print ("Accessing thread: %s" % fThread.components()[-1])

        # Create a copy of the thread file
        pThread.child("02 - thread.pdf").copy(temp.child("thread.pdf"))

        # Start assembling the new PDF file
        pdfFiles = []
        pdfFiles.append(temp.child("thread.pdf"))

        bookmarks = []
        bookmarks.append("Main Message Thread")

        # Collect the attachments and convert to PDF
        attachments = get_attachment_names(fThread)

        for attach in attachments:
            # Get original file details
            fFile = fThread.child(attach.file)
            ext = fFile.ext.upper()

            # Set up PDF details
            pdfName = "%s.pdf" % fFile.name
            pFile = temp.child(pdfName)

            if (ext == ".DOC" or ext == ".DOCX" or ext == ".DOT" 
                    or ext == ".DOCM" or ext == ".TXT" 
                    or ext == ".RTF" or ext == ".WPD"):
                # Create a title page
                titlePage = title_page(temp, attachment.title)
                pdfFiles.append(titlePage)

                # Add bookmark entry
                bookmarks.append(attachment.title)

                # Convert text documents to PDF
                pdfFiles.append(convert_word(fFile, pFile))

            elif (ext == ".PPT" or ext == ".PPTX"):
                # Create a title page
                titlePage = title_page(temp, attachment.title)
                pdfFiles.append(titlePage)

                # Add bookmark entry
                bookmarks.append(attachment.title)

                # Convert PowerPoint documents to PDF
                pdfFiles.append(convert_ppt())

            elif (ext == ".XLS" or ext == ".XLSX" or ext == ".XLTX" 
                  or ext == ".CSV"):
                # Create a title page
                titlePage = title_page(temp, attachment.title)
                pdfFiles.append(titlePage)

                # Add bookmark entry
                bookmarks.append(attachment.title)

                # Convert spreadsheet documents to PDF
                output = convert_xls()

                for file in output:
                    pdfFiles.append(file)
                
            elif (ext == ".PNG" or ext == ".GIF" or ext == ".JPG" 
                  or ext == ".JPEG"):
                # Create a title page
                titlePage = title_page(temp, attachment.title)

                # Add bookmark entry
                bookmarks.append(attachment.title)

                # Convert images to PDF
                pdfFiles.append(titlePage)
                pdfFiles.append(convert_image())

        # Generate the new final PDF name
        pdfName = "%s.pdf" % fThread.components()[-1]
        pdfLoc = pForum.child(pdfName)

        if len(pdfFiles) > 1:
            # Attachments present to merge
            merge = PdfFileMerger()

            for pdf in pdfFiles:
                merge.append(pdf)

            with open(pdfLoc, "rb") as mergeFile:
                merge.write(mergeFile)

            # add bookmarks

        else:
            # No attachment present to merge - rename file
            pdfFiles[0].rename(pdfLoc)

        # Clear temp folder
        for file in temp.listdir():
            file.remove()