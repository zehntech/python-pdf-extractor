import PyPDF2
import textract
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import xlsxwriter
import re
import json

workbook = xlsxwriter.Workbook('Test7.xlsx')
worksheet1 = workbook.add_worksheet("Drawing_List")
worksheet2 = workbook.add_worksheet("Equipment_List")

row = 1
column = 0

filename = '17d0011_Feed Preheat.C1_Red.pdf'

pdfFileObj = open(filename,'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
num_pages = pdfReader.numPages
# print(num_pages)
count = 0
text = ""
while count < num_pages:
    pageObj = pdfReader.getPage(count)
    count +=1
    text += pageObj.extractText()
if text != "":
   text = text
   # tokens = word_tokenize(text)
   text1 = text

   x = re.search(r"\d{2}-[A-Z]-\d{4}\s*\d+", text1)
   # print(x.group())
   data = x.group().split(" ")
   if(len(data[1]) == 1):
       data[1] = "0"+data[1]
   data[0] = data[0]+"Rev"+data[1];
   print(data[0])
   text2 = re.split("(\W+)\s*\d{2}-[A-Z]-\d{4}\s*\d+(.+)", text1)
   print(text2[0].split(". "))
   text3 = text2[0].split(". ")
   print('----------')
   print(text3[len(text3)-1])
   data.append("-")
   data.append(text3[len(text3)-1])
   print(data)
   worksheet1.write(0, 0, "Drawing Number/Rev")
   worksheet1.write(0, 1, "Revision")
   worksheet1.write(0, 2, "ACP Number")
   worksheet1.write(0, 3, "Diagram Name")

   for item in data:

       # if (item == "FUTURE"):

           worksheet1.write(row, column, item)
           # row += 1
           column += 1

   workbook.close()


else:
   text = textract.process("Legacy-Lea Unit West WC Battery-GL-102-REV.A.pdf", method='tesseract', language='eng')
























# import PyPDF2
# import textract
# from nltk.tokenize import word_tokenize
# from nltk.corpus import stopwords
# import xlsxwriter
#
# workbook = xlsxwriter.Workbook('Test.xlsx')
# worksheet = workbook.add_worksheet()
# row = 0
# column = 0
#
# filename = 'Legacy-Lea Unit West WC Battery-GL-102-REV.A.pdf'
#
# pdfFileObj = open(filename,'rb')
# pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
# num_pages = pdfReader.numPages
#
# count = 0
# text = ""
# while count < num_pages:
#     pageObj = pdfReader.getPage(count)
#     count +=1
#     text += pageObj.extractText()
# if text != "":
#    text = text
#    tokens = word_tokenize(text)
#    worksheet.write(0, 0, "Name")
#    for item in tokens:
#
#        if(item == "FUTURE") :
#            row += 1
#            worksheet.write(row, column, item)
#
#
#    workbook.close()
#
#
# else:
#    text = textract.process("Legacy-Lea Unit West WC Battery-GL-102-REV.A.pdf", method='tesseract', language='eng')
#
