import sys
import os
from unipath import Path, DIRS, FILES
from bs4 import BeautifulSoup
import pdfkit
import comtypes.client
from PyPDF2 import PdfFileMerger, PdfFileReader
from fpdf import FPDF
from PIL import Image

class AttachmentData:
    def __init__(self, url, title, file):
        self.url = url.strip()
        self.title = title.strip()
        self.file = file.strip()

def confirm_manual_conversions(psn):
    """Cycle through each directory and collect list of extensions"""
    # List of extensions that can be automatically converted
    autoExt = [".DOC", ".DOCX", ".DOT", ".DOTX", ".DOCM", ".TXT", 
               ".RTF", ".WPD", ".PPT", ".PPTX", ".XLS", ".XLSX",
               ".XLT", ".XLTX", ".CSV", ".PNG", ".GIF", ".JPG",
               ".JPEG", ".PDF", ".VCF"]

    # List any file that needs to be manually converted
    print ("Below are files that need to be manually converted to PDF:")

    manualConversion = False

    for forum in psn.listdir(filter=DIRS):
        for thread in forum.listdir(filter=DIRS):
            for file in thread.listdir(filter=FILES):
                if (file.ext.upper() not in autoExt 
                        and file.name != "thread.html" 
                        and file.name != "attachments.txt"):
                    print ("    %s" % file.absolute())
                    manualConversion = True

    if manualConversion:
        userResponse = input("\nHave all the above files been converted (Y/N)? ")
    
        if userResponse.upper() != "Y":
            print ("Please complete conversions and rerun program")
            sys.exit()

def get_style():
    style = "<style>"

    with open(root.child("style.css"), "r") as s:
        for line in s:
            style = "%s%s" % (style, line)

    style = "%s%s" % (style, "</style>")

    return style

def get_attachment_data(thread):
    aTxt = thread.child("attachments.txt")
    attachments = []

    if aTxt.exists():
        with open(aTxt, "r") as list:
            for a in list:
                split = a.split("    |    ")
                attachments.append(AttachmentData(split[0], split[1], split[2]))

    return attachments

def format_content(thread):
    oHTML = thread.child("thread.html")

    # Add document head
    nHTML = "<!DOCTYPE html><html><head>"
    nHTML = "%s%s" % (nHTML, "<meta charset='utf-8'>")
    nHTML = "%s%s" % (nHTML, style)
    nHTML = "%s%s" % (nHTML, "</head>")

    # Add the document body
    nHTML = "%s%s" % (nHTML, "<body>")

    with open(oHTML, encoding="utf-8", mode="r") as html:
        for line in html:
            nHTML = "%s%s" % (nHTML, line)

    nHTML = "%s%s" % (nHTML, "</body></html>")

    # Convert to a soup
    soup = BeautifulSoup(nHTML, "lxml")

    # Remove all links to the eForum (they will be inactive)
    links = soup.findAll("a", href=True)

    for link in links:
        if "index.cf" in link["href"] or "List.cfm" in link["href"] or \
        "post.cfm" in link["href"] or "print.cfm" in link["href"] or \
        "javascript:" in link["href"]:
            link.unwrap()

    # Convert the soup back to an HMTL string
    nHTML = soup.prettify(formatter="html")

    return nHTML

def set_attachment_links(html, type, thread, attachments=[]):
    if type == "html":
        # Change HTML links to relative links
        for attachment in attachments:
                # Correct for changes in HTML entities
                attachment.url = attachment.url.replace("&", "&amp;")
                attachment.file = attachment.file.strip()
                attachment.file = attachment.file.replace(" ", "%20")

                if attachment.url in html:
                    html = html.replace(attachment.url, attachment.file)
                else:
                    print ("Unable to update attachment %s" % attachment.title)
    else:
        # Remove attachment links for PDF
        soup = BeautifulSoup(html, "lxml")
        
        links = soup.findAll("a", href=True)
        
        for link in links:
            if "getAttachment.cfm" in link["href"]:
                link.unwrap()

        # Convert the soup back to an HMTL string
        html = soup.prettify(formatter="html")

    return html

def create_title_page(temp, title):
    # Format the title info
    title = Path(title).stem.strip()

    # Create a new PDF to hold the text
    pdf = FPDF("P", "in", "Letter")
    pdf.set_left_margin(1)
    pdf.set_top_margin(1)
    pdf.set_right_margin(1)
    pdf.add_page()

    # Set the font
    pdf.set_font("Courier", "B", 16)

    pdf.multi_cell(0, 0.5, title, 0, align="C")

    # Save the pdf
    outputFile = temp.child("title - %s.pdf" % title)
    pdf.output(outputFile, "F")

    return outputFile

def convert_word(inputFile, outputFile):
    """Converts: doc docx dot docm txt rtf wpd vcf"""
    # Open the Microsoft Word Application
    word = comtypes.client.CreateObject('Word.Application')

    # Open input file in Word
    doc = word.Documents.Open(inputFile.absolute().strip())

    # Save input file as PDF to output file
    doc.SaveAs(outputFile, FileFormat=17)

    # Close files and applications
    doc.Close()
    word.Quit()

def convert_ppt(inputFile, outputFile):
    """Converts: ppt pptx"""
    # Opens the Microsoft Powerpoint Application
    ppt = comtypes.client.CreateObject("Powerpoint.Application")

    # Open input file in Powerpoint
    deck = ppt.Presentations.Open(inputFile)

    # Save input file as PDF to output file
    deck.SaveAs(out_file, FileFormat=32)

    # Close files and applications
    deck.Close()
    ppt.Quit()

def convert_xls(inputFile, outputFile):
    """
        Converts: xls xlsx xltx csv
        Note: Saves each worksheet as separate file
    """
    # Opens the Microsoft Excel Application
    xls = comtypes.client.CreateObject("Excel.Application")
    
    # Open input file in Powerpoint
    wb = xls.Workbooks.Open(inputFile)
    
    # Cycle through each worksheet
    numWS = wb.Worksheets.Count
    fileNames = []

    for i in range(1, numWS+1):
        try:
            # Save worksheet as PDF to output file
            fileName = "%s_%s" % (outputFile.absolute(), i)
            fileNames.append(Path(fileName))
            wb.Worksheets[i].SaveAs(fileName, FileFormat=57)
        except Exception as e:
            None

    # Close files and applications
    wb.Close()
    xls.Quit()

    return fileNames

def convert_image(inputFile, outputFile):
    """Converts: png gif jpg jpeg"""
    # Open image and determine size
    image = Image.open(inputFile)
    height = image.size[0]
    width = image.size[1]
    
    # Create a new PDF to hold the image
    pdf = FPDF(unit="pt", format=[height, width])
    pdf.add_page()

    # Add the image to the page
    pdf.image(inputFile, 0, 0, 0, 0)

    # Save the pdf
    pdf.output(outputFile)

def format_html(root, thread, fForum, attachments):
    # Create a thread folder for HTML formatting
    fThread = fForum.child(thread.components()[-1])
    fThread.mkdir()

    # Take original thread and generate a new HTML strings
    fhtml = format_content(thread)

    # Formats attachment links properly
    html = set_attachment_links(fhtml, "html", thread, attachments)

    # Create the formatted HTML file
    with open(fThread.child("01 - thread.html"), encoding="utf-8", mode="w") as h:
        h.write(html)
            
    # Move over any attachments to the new folder
    for file in thread.listdir(filter=FILES):
        fileName= file.name.strip()

        if fileName != "thread.html" and fileName != "attachments.txt":
            print ("Copying %s" % fileName)
            file.copy(fThread.child(fileName))

def format_pdf(root, thread, fForum, temp, attachments):
    # Variables required to assemble final PDF
    pdfFiles = []
    bookmarks = []
    
    # Set up config for PDF generation
    path_wkthmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    pdfConfig = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf, )
    pdfOptions = {"quiet": ""}

    # Take original thread and generate a new HTML strings
    fhtml = format_content(thread)
 
    # Formats attachment links properly
    html = set_attachment_links(fhtml, "pdf", thread)

    # Convert the HTML file to a PDF
    pdfThread = temp.child("01 - thread.pdf")
    pdfkit.from_string(html, pdfThread, 
                        configuration=pdfConfig, options=pdfOptions)

    # Add thread PDF to list
    pdfFiles.append(pdfThread)
    bookmarks.append("Main Message Thread")

    for attachment in attachments:
        # Get original file details
        oFile = thread.child(attachment.file)
        ext = oFile.ext.strip().upper()
            
        # Set up PDF details
        pdfName = "%s.pdf" % oFile.stem.strip()
        fFile = temp.child(pdfName)

        if (ext == ".DOC" or ext == ".DOCX" or ext == ".DOT" 
                or ext == ".DOCM" or ext == ".TXT" 
                or ext == ".RTF" or ext == ".WPD" 
                or ext == ".VCF"):
            print ("Converting %s..." % attachment.title)

            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Convert text documents to PDF
            convert_word(oFile, fFile)

            # Add new PDF to file list
            pdfFiles.append(fFile)

        elif (ext == ".PPT" or ext == ".PPTX"):
            print ("Converting %s..." % attachment.title)

            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Convert PowerPoint documents to PDF
            convert_ppt(oFile, fFile)

            # Add new PDF to file list
            pdfFiles.append(fFile)

        elif (ext == ".XLS" or ext == ".XLSX" or ext == ".XLTX" 
                or ext == ".CSV"):
            print ("Converting %s..." % attachment.title)

            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Convert spreadsheet documents to PDF
            outputFiles = convert_xls(oFile, fFile)

            # Add new PDF files to file list
            for outputFile in outputFiles:
                pdfFiles.append(outputFile)
                
        elif (ext == ".PNG" or ext == ".GIF" or ext == ".JPG" 
                or ext == ".JPEG"):
            print ("Converting %s..." % attachment.title)

            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Convert images to PDF
            convert_image(oFile, fFile)

            # Add new PDF to file list
            pdfFiles.append(fFile)

        elif (ext == ".HTM" or ext == ".HTML" or ext == ".MHT"):
            print ("Converting %s..." % attachment.title)
            
            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Convert images to PDF
            convert_html(oFile, fFile)

            # Add new PDF to file list
            pdfFiles.append(fFile)

        elif (ext == ".PDF"):
            print ("Copying %s..." % attachment.title)

            # Create a title page
            titlePage = create_title_page(temp, attachment.title)
            pdfFiles.append(titlePage)

            # Add bookmark entry
            bookmarks.append(attachment.title)

            # Copy PDF to temp directory
            tempPDF = temp.child(oFile.name)
            oFile.copy(tempPDF)

            # Add copied PDF to file list
            pdfFiles.append(tempPDF)

    # Generate the new final PDF name
    pdfName = "%s.pdf" % thread.components()[-1]
    pdfLoc = fForum.child(pdfName)

    if len(pdfFiles) > 1:
        # Attachments present to merge
        merge = PdfFileMerger()

        for pdf in pdfFiles:
            pdf = PdfFileReader(pdf)

            if pdf.isEncrypted:
                try:
                    pdf.decrypt('')
                except:
                    pdf = root.child("encryptedPDF.pdf")
           
            merge.append(pdf)

        with open(pdfLoc, "wb") as mergeFile:
            merge.write(mergeFile)

        # add bookmarks

        merge.close()

    else:
        # No attachment present to merge - rename file
        pdfFiles[0].rename(pdfLoc)

    # Clear temp folder
    for file in temp.listdir():
        file.remove()

            
# APPLICATION SETUP
# Set up root path to generate absolute paths to files
formatType = sys.argv[1]
root = Path(sys.argv[2])

# Create the style sheet
style = get_style()

# Cycle through the PSN directories (main eForum directories)
print ("CSHP PSN Data Formatting Tool")
print ("----------------------------------------------------------------------")

psn = root.child("PSN")

if formatType == "html":
    fPSN = root.child("HTML Format")
elif formatType == "pdf":
    fPSN = root.child("PDF Format")
    confirm_manual_conversions(psn)


for forum in psn.listdir(filter=DIRS):
    print ("\nAccessing %s Forum" % forum.components()[-1])
    print ("----------------------------------------------------------------------")

    # Create the new forum folder
    fForum = fPSN.child(forum.components()[-1])
    fForum.mkdir()
    
    # Create a temp folder for PDF formatting
    if formatType == "pdf":
        temp = fForum.child("temp")
        temp.mkdir()

    for thread in forum.listdir(filter=DIRS):
        print ("Accessing thread: %s" % thread.components()[-1])

        attachments = get_attachment_data(thread)

        if formatType == "html":
            format_html(root, thread, fForum, attachments)
        elif formatType == "pdf":
            format_pdf(root, thread, fForum, temp, attachments)

    # Remove the temp folder
    if formatType == "pdf":
        temp.rmdir()

print ("\n----------------------------------------------------------------------")
print ("FORMATTING COMPLETE")
print ("----------------------------------------------------------------------\n")