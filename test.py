import gzip
from PIL import Image
from unipath import Path
from fpdf import FPDF

emz = Path("E:/Desktop/test.emz")

with gzip.open(emz, "rb") as emf:
    png = Image.open(emf).save("test.png")
    image = Image.open("test.png")
    outputFile = Path("E:/Desktop/test_image.pdf")

    """Converts: png gif jpg jpeg"""
    # Open image and determine size
    height = image.size[0]
    width = image.size[1]
    
    # Create a new PDF to hold the image
    pdf = FPDF(unit="pt", format=[height, width])
    pdf.add_page()

    # Add the image to the page
    pdf.image("test.png", 0, 0, 0, 0)

    # Save the pdf
    pdf.output(outputFile)