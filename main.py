# need to install module: PyPDF2, openpyxl, reportlab
# Read xlsx data and Write pdf
from PyPDF2 import PdfFileReader, PdfFileWriter
from openpyxl import load_workbook

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter

import io
import os
import re
import math

# add Font
pdfmetrics.registerFont(TTFont("D2Coding", "D2Coding.ttf"))

# filtering text
regex = re.compile(r'\d*[가-힣].*')
regex_count = re.compile(r'數量: \d+')

# get currentPath
currentPath = os.getcwd()
# get file list
files = os.listdir()
# filtering xlsx file
xlsx_files = list(filter(lambda s: True if s.find('.xlsx') > 0 else False, files))

for xlsx_file in xlsx_files:
    file_name = re.sub(".xlsx", "", xlsx_file)

    try:
        # Read pdf: Fix filename?
        exist_pdf = PdfFileReader(open(currentPath + '/' + file_name + '.pdf', "rb"))
        # Read xlsx Fix filename?
        wb = load_workbook(currentPath + '/' + file_name + '.xlsx')
        # sheet name
        sheet = wb['orders']
    except:                                     # if not exist pdf or xlsx
        continue

    # Write pdf
    output = PdfFileWriter()

    # get page num
    nums = exist_pdf.getNumPages()
    for num in range(nums):
        print(num)
        # col
        xlsx_num_data = "C{}".format(num + 2)
        str_list = str(sheet[xlsx_num_data].value).split('\n')
        count_list = list(map(regex_count.findall, str_list))
        str_list = list(map(regex.findall, str_list))
        str_list = list(map(lambda s: s[0], str_list))
        # make tuple to lists
        str_list = list(zip(count_list, str_list))

        posY = 75
        posY_chk = posY

        # create pdf Text
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("D2Coding", 7)  # set Font: D2Coding, Size: 7

        # posY 미리 계산기능 추가
        for idx, (count, str_data) in enumerate(str_list):  # drawing Text, idx, str_data, exception case: G01305644544
            if posY_chk < 0:
                posY_chk = posY_chk + 10
                posY = posY + 10
            texts = re.sub(".*:|;.*", "", str_data).strip().split(',')
            text_len = math.ceil(len(texts) / 3)
            if text_len > 1:
                for i in range(text_len):
                    posY_chk -= 10
            else:
                posY_chk -= 10

        for idx, (count, str_data) in enumerate(str_list):  # drawing Text, idx, str_data, exception case: G01305644544
            texts = re.sub(".*:|;.*", "", str_data).strip().split(',')
            counts = re.sub('數量: ', " ", count[0])
            text_len = math.ceil(len(texts) / 3)
            if text_len > 1:
                for i in range(text_len):
                    startPoint = i * 3
                    endPoint = (i + 1) * 3
                    write_text = ",".join(texts[startPoint:endPoint])
                    if i is text_len - 1:
                        write_text = write_text + counts
                    can.drawString(65, posY, write_text)
                    posY -= 10
            else:
                can.drawString(65, posY, ",".join(texts) + counts)
                posY -= 10
        can.save()
        packet.seek(0)
        new_pdf = PdfFileReader(packet)

        # get Page
        page = exist_pdf.getPage(num)
        # page Merging
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)

    # write data to create.pdf
    with open(file_name+"_create.pdf", "wb") as outputStream:
        output.write(outputStream)


# https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python
