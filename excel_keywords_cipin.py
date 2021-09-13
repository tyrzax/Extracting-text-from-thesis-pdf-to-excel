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
        #选择想要读取哪一列的文本
            cell_value1 = table.cell(row, 5).value
            #cell_value2 = table.cell(row, 6).value
            #cell_value3 = table.cell(row, 7).value
            # 把数据追加到excel_list中
            excel_list.append(cell_value1)
            #excel_list.append(cell_value2)
            #excel_list.append(cell_value3)
        print(excel_list)
        return excel_list
    except Exception:
        print('没有')

xls = '/Users/tyrzax/ZJU Documents/all__六大期刊合并(23164).xlsx'
name = xls.split('.')[0]

textlist = excel_data(xls)
text = ';'.join(textlist)
stopwords = [line.strip() for line in open('/Users/tyrzax/ZJU Documents/stopwords.txt').readlines()]
cixing = pseg.lcut(text)
count = jieba.lcut(text)
word_count = {}
word_flag = {}
all = []

with codecs.open(filename= name+'关键词.csv', mode='w')as f:
    write = csv.writer(f, dialect='excel')
    write.writerow(["word", "count", "flag"])
    # 词性统计
    for w in cixing:
        word_flag[w.word] = w.flag

    # 词频统计
    for word in count:
        if word not in stopwords:
            # 不统计字数为一的词
            if len(word) == 1:
                continue
            else:
                word_count[word] = word_count.get(word, 0) + 1

    items = list(word_count.items())
    # 按词频排序
    items.sort(key=lambda x: x[1], reverse=True)

    for i in range(len(items)):
        word = []
        word.append(items[i][0])
        word.append(items[i][1])
        # 若词频字典里，该关键字有分辨出词性，则记录，否则为空
        if items[i][0] in word_flag.keys():
            word.append(word_flag[items[i][0]])
        else:
            word.append("")
        all.append(word)

    for res in all:
        write.writerow(res)

