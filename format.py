import sys
from unipath import Path, DIRS
from bs4 import BeautifulSoup
import pdfkit

def get_style():
    style = "<style>"

    with open(root.child("style.css"), "r") as s:
        for line in s:
            style = "%s%s" % (style, line)

    style = "%s%s" % (style, "</style>")

    return style

def format_content(thread):
    oHTML = thread.child("thread.html")

    # Add document head
    nHTML = "<!DOCTYPE html><html><head><meta charset='utf-8'>"
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

    # Update the links to attachments to be relative links

    return soup.prettify()

# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

# Set up config for PDF generation
path_wkthmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
pdfConfig = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
pdfOptions = {"quiet": ""}

# Create the style sheet
style = get_style()

# Cycle through the PSN directories (main eForum directories)
psn = root.child("PSN")
fPSN = root.child("formattedPSN")

for forum in psn.listdir(filter=DIRS):
    print ("\nAccessing %s Forum" % forum.components()[-1])
    print ("----------------------------------------------------------------------")

    # Create the new forum folder
    fForum = fPSN.child(forum.components()[-1])
    fForum.mkdir()

    for thread in forum.listdir(filter=DIRS):
        print ("Accessing thread: %s" % thread.components()[-1])

        # Create the new thread folder
        fThread = fForum.child(thread.components()[-1])
        fThread.mkdir()

        # Take original thread and generate a new HTML string
        html = format_content(thread)
        
        # Create the new html document for backup
        with open(fThread.child("thread.html"), encoding="utf-8", mode="w") as h:
            h.write(html)
        
        # Create the new PDF document 
        pdfkit.from_string(html, fThread.child("thread.pdf"), 
                           configuration=pdfConfig, options=pdfOptions)

        # Move over any attachments to the new folder