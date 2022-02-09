from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas

output = PdfFileWriter()

with open("test.pdf", "wb") as outputStream:
    output.write(outputStream)