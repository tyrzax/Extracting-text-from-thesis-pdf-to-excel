from io import StringIO, open
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from openpyxl import load_workbook
import xlwt
import os
import regex as re
import jieba.posseg as pseg

def convert(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = open(fname,'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text

def convertMultiple(pdfDir, txtDir):
    if pdfDir == "": pdfDir = os.getcwd() + "//"
    for pdf in os.listdir(pdfDir):
        fileExtension = pdf.split(".")[-1]
        if fileExtension == "pdf":
            pdfFilename = pdfDir + pdf
            text = convert(pdfFilename)
            textFilename = txtDir + pdf + ".txt"
            textFile = open(textFilename, "w",encoding = 'utf-8')
            textFile.write(text)
            textFile.close()

pdfDir = "C://Users/GUOXIAO/Desktop/yes/pdf/"
txtDir = "C://Users/GUOXIAO/Desktop/yes/txt/"
excelDir = "C://Users/GUOXIAO/Desktop/yes/excel/"
convertMultiple(pdfDir, txtDir)
print("Extraction completed.")


if txtDir == "": txtDir = os.getcwd() + "//"
for txt in os.listdir(txtDir):
    fileExtension = txt.split(".")[-1]
    if fileExtension == "txt":
        txtFilename = txtDir + txt
        textFile = open(txtFilename,encoding='utf-8')
        raw=textFile.read()
        textFile.close()
        raw4=raw.replace("\n","").replace("\r","").replace("(","（").replace(")","）").replace("\u2003","").replace("Tile:","Title:")
        raw2=raw4.split("")
        raw3="".join(raw2)
        raw5=raw.replace(")","）").replace("\n","")
        
        keywords=[]
        
        for line in raw.split("\n"):
            if "关键词：" in line:
                keywords.append(line.replace("关键词：",""))
        keywords.remove(keywords[0])
            
        cntitle= re.findall('(?<=）).*?(?=…………)',raw5)
        firstAuthor= re.findall('(?<=作者简介：).*?(?=，)',raw3)
        engtitle= re.findall('(?<=Title:).*?(?=Abstract:)',raw3)
        abstract = re.findall('(?<=内容摘要：).*?(?=关键词：)',raw3)
        programs= re.findall('(?<=基金项目：).*?(?=作者简介：)',raw3)
        abstract.remove(abstract[0])
        firstAuthor.remove(firstAuthor[0])
        cntitle.remove(cntitle[-1])

        rawAuthorInfo= re.findall('(?<=关键词：).*?(?=Title:)',raw3)
        authorInfos="\n".join(rawAuthorInfo)
        secondAuthor=[]
        SA= re.findall('(?<=；).*?(?=，)',authorInfos)
        for names in SA:
            if len(names)<10:
                secondAuthor.append(names)
        iWantUniversities=authorInfos.split("，")

        university=[]
        for item in iWantUniversities:
            if "大学" in item:
                university.append(item)
            elif "学院" in item:
                university.append(item)
        university.remove(university[-1])
        university.remove(university[1])
        university.remove(university[0])


        #写入表格
        wbk = xlwt.Workbook(encoding='ascii')
        sheet = wbk.add_sheet("info")
        i=0
        j=0
        k=0
        l=0
        m=0
        n=0
        o=0
        p=0
        for a in cntitle:
            sheet.write(i, 0,a)
            i=i+1
        for b in engtitle:
            sheet.write(j, 1,b)
            j=j+1
        for c in firstAuthor:
            sheet.write(k, 2,c)
            k=k+1
        for d in university:
            sheet.write(l, 3,d)
            l=l+1
        for e in secondAuthor:
            sheet.write(m, 6,e)
            m=m+1
        for f in keywords:
            sheet.write(n, 10,f)
            n=n+1
        for g in abstract:
            sheet.write(o, 11,g)
            o=o+1
        for h in programs:
            sheet.write(p, 12,h)
            p=p+1
        wbk.save(txtFilename+'infos.xls')  

print("Done.")


        


            

