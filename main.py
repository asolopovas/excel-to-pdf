import io
import os
from win32com import client
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdfCropMargins import crop


def delete_files(files):
    for file in files:
        if os.path.exists(file):
            os.remove(file)

delete_files(['template_resize.pdf', 'template_cropped.pdf'])

# Open Microsoft Excel
excel = client.Dispatch("Excel.Application")

# Read Excel File
sheets = excel.Workbooks.Open('C:\\Users\\asolo\\git\\excel-to-pdf\\template.xlsx')
work_sheets = sheets.Worksheets[0]

# Remove all footers and headers
work_sheets.PageSetup.LeftFooter = ""
work_sheets.PageSetup.CenterFooter = ""
work_sheets.PageSetup.RightFooter = ""
work_sheets.PageSetup.LeftHeader = ""
work_sheets.PageSetup.CenterHeader = ""
work_sheets.PageSetup.RightHeader = ""

# Convert into PDF File
work_sheets.ExportAsFixedFormat(0, 'C:\\Users\\asolo\\git\\excel-to-pdf\\template.pdf')

# close document without saving
sheets.Close(False)
excel.Application.Quit()

crop(["-p", "0", "template.pdf"])

def resize_pdf(input_pdf, width):
    file = open(input_pdf, 'rb')
    pdf = PdfFileReader(file)
    page0 = pdf.getPage(0)
    content_width = page0.mediaBox.getWidth()
    content_height = page0.mediaBox.getHeight()

    # convert width to mm
    content_width_mm=float(content_width)/72.0*25.4

    # coluate sacle factor
    scale_factor = width/content_width_mm

    page0.scaleBy(scale_factor)  # float representing scale factor - this happens in-place
    # close the input file

    return page0

page = resize_pdf("template_cropped.pdf", 100)


# create letter size pdf and place page into it
packet = io.BytesIO()
can = canvas.Canvas(packet, pagesize=letter)
can.setPageSize((page.mediaBox.getWidth(), page.mediaBox.getHeight()))
can.doForm('page', page)
can.save()



