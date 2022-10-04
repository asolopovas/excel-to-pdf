from calendar import c
import os
import tqdm
import tempfile
from win32com import client
from PyPDF2 import PdfMerger
from pdfCropMargins import crop
from lib import delete_files
from PyPDF2 import PdfFileReader, PdfFileWriter

tmp = os.path.join(tempfile.gettempdir(), 'lyntouch-excel-to-pdf')
if not os.path.exists(tmp):
    os.makedirs(tmp)


def excelToPdf(src, output):
    excel = client.Dispatch("Excel.Application")

    sheets = excel.Workbooks.Open(src)
    page = 1
    for sheet in sheets.Sheets:
        tmpPath = os.path.join(tmp, str(page) + ".pdf")
        sheet.PageSetup.LeftFooter = ""
        sheet.PageSetup.CenterFooter = ""
        sheet.PageSetup.RightFooter = ""
        sheet.PageSetup.LeftHeader = ""
        sheet.PageSetup.CenterHeader = ""
        sheet.PageSetup.RightHeader = ""
        sheet.ExportAsFixedFormat(0, tmpPath)
        page += 1
        trimPDFMargins(tmpPath)

    mergePDFs(tmp, output)

    sheets.Close(False)
    excel.Application.Quit()


def mergePDFs(src, output):
    for file in os.listdir(src):
        merger = PdfMerger()
        if file.endswith(".pdf"):
            merger.append(os.path.join(src, file))
        merger.write(output)
        merger.close()


def trimPDFMargins(src):
    # get filename without extension
    filename = os.path.splitext(src)[0]
    if os.path.exists(filename + "_cropped.pdf"):
        delete_files([filename + "_cropped.pdf"])
    crop(["-p", "0", os.path.join(tmp, src), "-o", os.path.join(tmp, filename + "_cropped.pdf")])
    delete_files([src])
    os.rename(filename + "_cropped.pdf", src)
    addPDFMargins(src)

def addPDFMargins(src):
    with open(src, 'rb') as f:
        p = PdfFileReader(f)
        info = p.getDocumentInfo()
        number_of_pages = p.getNumPages()

        writer = PdfFileWriter()
        margin = 30
        print(f'margin: {margin}')
        print (info)
        print (f'info: {info}')
        print (f'number_of_pages: {number_of_pages}')
        for i in range(number_of_pages):
            print( i)
            # page = p.getPage(i)
            # new_page = writer.addBlankPage(
            #     page.mediaBox.getWidth() + 2 * margin,
            #     page.mediaBox.getHeight() + 2 * margin
            # )
            # new_page.mergeScaledTranslatedPage(page, 1, margin, margin)
            # # writer.addPage(new_page)
        with open(src, 'wb') as f:
            writer.write(f)



