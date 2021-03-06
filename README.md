# CSHP PSN Extraction Tool
Tool to automate archiving of the CSHP PSN eForums

# Background
This tool is used to scrape all data and attachments from the CSHP 
eForum. Access to the eForum is/was required for this to program to 
work. This required connecting to the eForum and passing the login 
credentials to the script, iterating over the entire forum and 
downloading the content and any attachments. These files were then 
processed to either allow viewing in a web browser locally or were 
all converted into a single PDF. Given that this was a one-time 
project with a deadline, best-practices were attempted, but often
not used (because in the end, this script only has to work once!)


# Useful Content
The following is unique/useful content that could be applicable projects:
 - Logging into a web-based system and passing the credentials to a 
   Python script
    - Uses the selenium webdriver library
 - Navigating through web pages with a Python script
    - Uses the requests library
 - Conversion of multiple file types to PDF (.doc, .docx, .dot, .dotx, 
   .docm, .txt, .rtf, .wpd, .ppt, .pptx, .xls, .xlsx, .xlt, xltx, .csv, 
   .png, .gif, .jpg/jpeg, .html/htm, .emz, .bmp)
    - uses the comtypes, PIL, gzip, pdfkit, and fpdf libraries
 - Merging multiple PDFs into a single PDF
    - Uses the PyPDF2 library