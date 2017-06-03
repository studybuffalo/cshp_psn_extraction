import sys
from unipath import Path, DIRS, FILES
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

def set_attachment_links(html, type, thread):
    # Update the links to attachments to be relative links
     # Update the links to attachments to be relative links
    aTxt = thread.child("attachments.txt")
    aList = []

    if aTxt.exists():
        with open(aTxt, "r") as list:
            for a in list:
                split = a.split("    |    ")
                
                # Correct for changes in HTML entities
                split[0] = split[0].replace("&", "&amp;")
                split[2] = split[2].strip()
                split[2] = split[2].replace(" ", "%20")

                if type == "pdf":
                    split[2] = "file:///%s" % split[2]

                if (split[0] in html):
                    html = html.replace(split[0], split[2])
                else:
                    print ("Unable to update attachment %s" % split[1])

    return html

# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

# Set up config for PDF generation
path_wkthmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
pdfConfig = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf, )
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

        # Take original thread and generate a new HTML strings
        fhtml = format_content(thread)
        html = set_attachment_links(fhtml, "html", thread)
        pdf = set_attachment_links(fhtml, "pdf", thread)

        # Create the new html document for backup
        with open(fThread.child("01 - thread.html"), encoding="utf-8", mode="w") as h:
            h.write(html)
        
        # Create the new PDF document 
        pdfkit.from_string(pdf, fThread.child("02 - thread.pdf"), 
                           configuration=pdfConfig, options=pdfOptions)

        # Move over any attachments to the new folder
        for attachment in thread.listdir(filter=FILES):
            print (attachment.name)
            if attachment.name != "attachments.txt" and \
               attachment.name != "thread.html":
                attachment.copy(fThread.child(attachment.name))