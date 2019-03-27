import PyPDF2
import textract
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import xlsxwriter
import re
import json
import argparse
import os
import sys

# from argparse import ArgumentParser
# parser = ArgumentParser()
# parser.add_argument("-f", "--file", dest="myFile", help="Open specified file")
# args = parser.parse_args()
# myFile = args.myFile
# filename = myFile
# print(filename)
print(sys.argv[1])
print(sys.argv[2])

# This is the one part I added (the read() call)
# text = open(myFile)
# print(text.read())
filename = sys.argv[1]
output = sys.argv[2]

workbook = xlsxwriter.Workbook(output)
worksheet1 = workbook.add_worksheet("Drawing_List")
worksheet2 = workbook.add_worksheet("Equipment_List")

row = 1
column = 0

# filename = '17d0011_Feed Preheat.C1_Red.pdf'

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
   print(text)
   # tokens = word_tokenize(text)
   print(text)
   text1 = text

   x = re.search(r"\d{2}-[A-Z]-\d{4}\s*\d+", text1)
   # print(x.group())
   data = x.group().split(" ")
   text2 = re.split("(\W+)\s*\d{2}-[A-Z]-\d{4}\s*\d+(.+)", text1)
   # print(text2[0].split(". "))
   text3 = text2[0].split(". ")
   # print('----------')
   print(text3[len(text3)-1])
   data.append("-")
   data.append(text3[len(text3)-1])
   print(data)

   worksheet1.write(0, 0, "Drawing Number/Rev")
   worksheet1.write(0, 1, "Revision")
   worksheet1.write(0, 2, "ACP Number")
   worksheet1.write(0, 3, "Diagram Name")

   for item in data:
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
