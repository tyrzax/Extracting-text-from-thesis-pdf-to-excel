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
    txt = text
    cixing = pseg.lcut(txt)
    count = jieba.lcut(txt)
    word_count = {}
    word_flag = {}
    all = []

    with codecs.open('/Users/tyrzax/ZJU Documents/六大期刊.csv','w')as f:
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

xls_path = '/Users/tyrzax/ZJU Documents/all__六大期刊合并(23164).xlsx'


df = pd.read_excel(xls_path)
df_1 = df[df[''].str.contains('')]
df_1 = df_1.drop(columns=['英文题名','第一作者（含职称职务等信息）','机构','第二作者（含职称职务等信息）','机构','项目','期刊','国别','国别词频参考'])
print(df_1)




yanjiuwenti(text)
