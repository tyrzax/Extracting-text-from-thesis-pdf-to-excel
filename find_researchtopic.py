# -*- coding: utf-8 -*-
import jieba
import jieba.posseg as pseg
import xlrd
import pandas as pd
import codecs

def yanjiuwenti(text,filepath,infocount):
    stopwords = [line.strip() for line in
                 open('/Users/tyrzax/ZJU Documents/stopwords.txt').readlines()]
    txt = text
    count = jieba.lcut(txt)
    word_count = {}
    for word in count:
        if word not in stopwords:
            if len(word) == 1:
                continue
            else:
                word_count[word] = word_count.get(word, 0) + 1

    #print(word_count)
    items = list(word_count.items())
    items.sort(key=lambda x: x[1], reverse=True)
    item = dict(items)
    K_words = [key for key, value in item.items()]
    print(item)
    return K_words

def excel_data(file):
    name = file.split('.')[0]
    # 打开Excel文件读取数据
    data = xlrd.open_workbook(file)
    # 获取第一个工作表
    table = data.sheet_by_index(0)
    # 获取行数
    nrows = table.nrows

    info_list = []
    info_count= 0

    for row in range(1, nrows):
        cell_value1 = table.cell(row, 1).value
        cell_value2 = table.cell(row, 5).value
        cell_value3 = table.cell(row, 10).value
        # 把数据追加到excel_list中
        info = cell_value1 + '\n' + cell_value2
        if '帕斯捷尔纳克' in info:
            info_list.append(info)
            if int(cell_value3) > 2010:
                info_count = info_count + 1
    info = '\n'.join(info_list)
    print(info_count)
    #yanjiuwenti(info1,file,info_count1)
    word_list = yanjiuwenti(info,file,info_count)
    word_count_list = []
    for word in word_list:
        counter = 0
        for info in info_list:
            if word in info:
                counter = counter+1
        word_count_list.append(counter)
    dictionary = dict(zip(word_list,word_count_list))
    print(dictionary)
    filename = name+'研究问题帕斯捷尔纳克_'+str(info_count)+'.xlsx'

    df = pd.DataFrame.from_dict(dictionary, orient='index', columns=['count'])
    df = df.reset_index().rename(columns={'index': 'word'})
    writer = pd.ExcelWriter(filename)
    df.to_excel(filename)


excel_data('/Users/tyrzax/ZJU Documents/近20年整理/当代外国文学2000-2020.xlsx')