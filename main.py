import inspect
import io
import os
import sys
import tempfile
import argparse
from win32com import client
from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdfCropMargins import crop
from lib import delete_files

parser = argparse.ArgumentParser()
parser.add_argument('--foo', help='foo help')
args = parser.parse_args()

dir_tmp = tempfile.gettempdir()
dir_exec = os.getcwd()
dir_script = os.path.dirname(
    os.path.abspath(inspect.getfile(inspect.currentframe()))
)

# file_src = os.path.abspath(sys.argv[1])
# if not file_src.endswith(".xlsx"):
#     print("Error: File extension must be .xlsx")
#     exit()
# filename = os.path.splitext(file_src)[0]

# file_out = filename + ".pdf"
# # if output file exist
# if os.path.exists(file_out) or os.path.exists(filename + "_cropped.pdf"):
#     delete_files([file_out, filename + "_cropped.pdf"])

# file_outPath = os.path.join(dir_exec, file_out)

# # Open Microsoft Excel
# excel = client.Dispatch("Excel.Application")

# # Read Excel File
# sheets = excel.Workbooks.Open(file_src)
# work_sheets = sheets.Worksheets[0]

# # Remove all footers and headers
# work_sheets.PageSetup.LeftFooter = ""
# work_sheets.PageSetup.CenterFooter = ""
# work_sheets.PageSetup.RightFooter = ""
# work_sheets.PageSetup.LeftHeader = ""
# work_sheets.PageSetup.CenterHeader = ""
# work_sheets.PageSetup.RightHeader = ""

# work_sheets.ExportAsFixedFormat(0, file_outPath)

# # close document without saving
# sheets.Close(False)
# excel.Application.Quit()

# crop(["-p", "0", file_outPath])

# # def resize_pdf(input_pdf, width):
# #     file = open(input_pdf, 'rb')
# #     pdf = PdfFileReader(file)
# #     page0 = pdf.getPage(0)
# #     content_width = page0.mediaBox.getWidth()
# #     content_height = page0.mediaBox.getHeight()

# #     # convert width to mm
# #     content_width_mm=float(content_width)/72.0*25.4

# #     # coluate sacle factor
# #     scale_factor = width/content_width_mm

# #     page0.scaleBy(scale_factor)  # float representing scale factor - this happens in-place
# #     # close the input file

# #     return page0

# # page = resize_pdf("template_cropped.pdf", 100)
