# -*- coding: utf-8 -*-
import os
import xlrd
import xlwt
import re
from collections import Counter
import pandas as pd
import jieba
import jieba.posseg as pseg
import csv
import codecs

def yanjiuwenti(text):
    stopwords = [line.strip() for line in
                 open('/Users/tyrzax/ZJU Documents/stopwords.txt').readlines()]
    t = open(text,'r')
    txt = t.read()
    cixing = pseg.lcut(txt)
    count = jieba.lcut(txt)
    word_count = {}
    word_flag = {}
    all = []

    with codecs.open(text+'.csv','w')as f:
        write = csv.writer(f,'excel')
        write.writerow(["word", "count", "flag"])

        for w in cixing:
            word_flag[w.word] = w.flag

        for word in count:
            if word not in stopwords:
                if len(word) == 1:
                    continue
                else:
                    word_count[word] = word_count.get(word, 0) + 1
        word_count = {key:val for key, val in word_count.items() if val != 1}
        print(word_count)
        items = list(word_count.items())
        items.sort(key=lambda x: x[1], reverse=True)
        for i in range(len(items)):
            word = []
            word.append(items[i][0])
            word.append(items[i][1])
            if items[i][0] in word_flag.keys():
                word.append(word_flag[items[i][0]])
            else:
                word.append("")
            all.append(word)

        for res in all:
            write.writerow(res)


def excel_data(file):
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(file)
        # 获取第一个工作表
        table = data.sheet_by_index(0)
        # 获取行数
        nrows = table.nrows
        # 获取列数
        ncols = table.ncols
        # 定义excel_list
        excel_list = []
        for row in range(1, nrows):
            cell_value1 = table.cell(row, 5).value
            #cell_value2 = table.cell(row, 6).value
            #cell_value3 = table.cell(row, 7).value
            # 把数据追加到excel_list中
            excel_list.append(cell_value1)
            #excel_list.append(cell_value2)
            #excel_list.append(cell_value3)
        return excel_list
    except Exception:
        print('没有')

xls_list = ['/Users/tyrzax/ZJU Documents/all__六大期刊合并(23164).xlsx']

sheets = []
for excel in xls_list:
    list = excel_data(excel)
    text = '\n'.join(list)
    bookTitle = re.findall(".*《(.*)》.*",text)
    bookCount = Counter(bookTitle)
    count = dict(bookCount)
    df = pd.DataFrame(count.items(),columns = ['title','count'])
    name = excel.split('.')[0]
    df.sort_values(by='count', ascending=False).to_excel(name+'关键词.xls')


