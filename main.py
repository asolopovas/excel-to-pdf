import inspect
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdfCropMargins import crop
from lib import delete_files
from lib.config import getConfig
from lib.pdf import excelToPdf

conf = getConfig()
excelToPdf(conf.input, conf.output)


# if os.path.exists(file_out) or os.path.exists(filename + "_cropped.pdf"):
#     delete_files([file_out, filename + "_cropped.pdf"])
# crop(["-p", "0", file_outPath])

# # # def resize_pdf(input_pdf, width):
# # #     file = open(input_pdf, 'rb')
# # #     pdf = PdfFileReader(file)
# # #     page0 = pdf.getPage(0)
# # #     content_width = page0.mediaBox.getWidth()
# # #     content_height = page0.mediaBox.getHeight()

# # #     # convert width to mm
# # #     content_width_mm=float(content_width)/72.0*25.4

# # #     # coluate sacle factor
# # #     scale_factor = width/content_width_mm

# # #     page0.scaleBy(scale_factor)  # float representing scale factor - this happens in-place
# # #     # close the input file

# # #     return page0

# # # page = resize_pdf("template_cropped.pdf", 100)
