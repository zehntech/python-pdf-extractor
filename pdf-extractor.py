import PyPDF2
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import xlsxwriter
import re
import json
import argparse
import os
import sys
from pprint import pprint

def walk(obj, fnt, emb):
    
    if not hasattr(obj, 'keys'):
        return None, None
    fontkeys = set(['/FontFile', '/FontFile2', '/FontFile3'])
    if '/BaseFont' in obj:
        fnt.add(obj['/BaseFont'])
    if '/FontName' in obj:
        if [x for x in fontkeys if x in obj]:# test to see if there is FontFile
            emb.add(obj['/FontName'])

    for k in obj.keys():
        walk(obj[k], fnt, emb)

    return fnt, emb# return the sets for each page

def init():
    parser = argparse.ArgumentParser()
    parser.add_argument('input')
    parser.add_argument('output')
    args = parser.parse_args()
    drawing_list_tab = [{'key': 'drawing_number_rev', 'label': 'Drawing Number/Rev'}, {'key': 'revision', 'label': 'Revision'}, {'key': 'apc_number', 'label': 'APC Number'},
                        {'key': 'diagram_name', 'label': 'Diagram Name'}]
    equipment_list_tab = [{'key': 'drawing_number_rev', 'label': 'Drawing Number/Rev'}, {'key': 'type', 'label': 'Type'}, {'key': 'tag_number', 'label': 'Tag Number'},
                          {'key': 'equipment_full_name', 'label': 'Equipment Full Name'}, {
                              'key': 'section', 'label': 'Section'}, {'key': 'design', 'label': 'Design'},
                          {'key': 'remarks', 'label': 'Remarks'}]
                     
    print("input" + args.input)
    print("output "+args.output)
   
    filename = args.input
    output = args.output

    workbook = xlsxwriter.Workbook(output)
    worksheet1 = workbook.add_worksheet("Drawing List")
    worksheet2 = workbook.add_worksheet("Equipment List")

    row = 1
    column = 0

    # filename = '17d0011_Feed Preheat.C1_Red.pdf'

    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    num_pages = pdfReader.numPages
    print(num_pages)
    count = 0
    text = ""
    while count < num_pages:
        pageObj = pdfReader.getPage(count)
        print("txt "+pageObj.extractText())
        count += 1
        text += pageObj.extractText()
    if text != "":
        text = text
        # print(text)
        # tokens = word_tokenize(text)
        #print(text)
        text1 = text
        mfile = open('.\m-test.txt', 'wb')
        mfile.write(text)    
        print(text)
        return
        # txtFile=open('./output/txtFile.txt',"w",encoding='utf-8')
        # txtFile.write(text1)
        x = re.search(r"\d{2}-[A-Z]-\d{4}\s*\d+", text1)
        # print(x.group())
        data = {'drawing_number_rev': x.group().split(
            " ")[0], 'revision': x.group().split(" ")[1]}
        text2 = re.split("(\W+)\s*\d{2}-[A-Z]-\d{4}\s*\d+(.+)", text1)
        # print(text2[0].split(". "))
        text3 = text2[0].split(". ")
        # print('----------')
        #print(text3[len(text3)-1])
        data['apc_number'] = '-'
        data['diagram_name'] = text3[len(text3)-1]
        #print(data)
        row = 0
        col = 0
        for item in drawing_list_tab:
            worksheet1.write(row, col, item['label'])
            col += 1

        row = 0
        col = 0
        for item in equipment_list_tab:
            worksheet2.write(row, col, item['label'])
            col += 1
        row = 1
        col = 0
        for key in data:
            value = data[key]
            if(key == 'drawing_number_rev'):
                value = data['drawing_number_rev']+'Rev '+data['revision'] if int(
                    data['revision']) > 9 else data['drawing_number_rev']+'Rev 0'+data['revision']
            worksheet1.write(row, column, value)
            column += 1
        equipment_list = get_equipment_list(text1, data)
        row = 1
        col = 0
        for equipment in equipment_list:
            col=0
            for key in equipment:
                value = equipment[key]
                if(key=='revision'):
                    continue
                if(key == 'drawing_number_rev'):
                    value = equipment['drawing_number_rev']+'Rev '+equipment['revision'] if int(
                        equipment['revision']) > 9 else equipment['drawing_number_rev']+'Rev 0'+equipment['revision']
                worksheet2.write(row, col, value)
                col +=1
            row += 1
        workbook.close()


def get_equipment_list(text, data):
    equipment_list = []
    equipment_name_regex=["ARC [lI] VA\w+ RE\w+D","VAC\w+M \w+NIT"]
    equipment_type_regex=["exchanger","feed pumps"]    
    remaining = text
    matches=[]
    while  (re.search(r"\s+([A-Z]-\d{4}[A-Z\d](/[A-Z])?)\s+[^,]", remaining) is not None):
        match = re.search(r"\s+([A-Z]-\d{4}[A-Z\d](/[A-Z])?)\s+[^,]", remaining)
        print(match.group(1))
        matches.append({'tag_number':match.group(1)})
        remaining = remaining[match.end(1):] 

    remaining = text
    for index, match in enumerate(matches):
        for item in equipment_type_regex:
            if re.search(item,remaining,re.IGNORECASE) is not None:
                match['equipment_type']=re.search(item,remaining,re.IGNORECASE).group()
                #print(remaining)
                print(re.search(item,remaining,re.IGNORECASE).group())
                start=re.search(item,remaining,re.IGNORECASE).start()
                start +=len(re.search(item,remaining,re.IGNORECASE).group()) 
                remaining = remaining[start:len(remaining)] 
                break 
    
    remaining = text
    for match in matches:
        for item in equipment_name_regex:
            if re.search(item,remaining) is not None:
                #print(remaining)
                match['equipment_name']=re.search(item,remaining,re.IGNORECASE).group()
                print(re.search(item,remaining).group())
                start=re.search(item,remaining).start()
                start +=len(re.search(item,remaining).group()) 
                remaining = remaining[start:len(remaining)] 
                break 

        

    pprint(matches)
    for match in matches:
        equipment_list.append({
            'drawing_number_rev': data['drawing_number_rev'],
            'revision': data['revision'],
            'type':'-',
            'tag_number':match[0]

        })
    return equipment_list
    #pprint(match[0])


if __name__ == '__main__':
    init()
