# -*- coding: utf-8 -*-
import os, sys
import xlrd
import xlwt
from xlutils.copy import copy
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import re
import numpy as np
import csv
import codecs
from xlutils.copy import copy



def getResearchedAuthors(text):
    author_japan = ['川端康成','芥川龙之介','夏目漱石','三岛由纪夫','村上春树','大江健三郎','太宰治','松本清张','渡边淳一']
    author_india = ['泰戈尔']
    author_belgium = ['梅特林克','维尔哈伦','高斯特']
    author_france = ['卢梭','罗曼·罗兰','雨果','巴尔扎克','高乃依','大仲马','小仲马','莫泊桑','萨特','梅里美','福楼拜','杜拉斯','左拉','乔治·桑','普鲁斯特','勒克莱齐奥']
    author_uk = ['莎士比亚','弥尔顿','简·奥斯汀','狄更斯','济慈','王尔德','笛福','彭斯','萧伯纳','奥威尔','弗吉尼亚·伍尔夫','阿加莎·克里斯蒂','J.K.罗琳']
    author_belarus = ['斯韦特兰娜·阿列克西耶维奇']
    author_russia = ['普希金','莱蒙托夫','果戈里','果戈理','屠格涅夫','涅克拉索夫','丘特切夫','托尔斯泰','陀思妥耶夫斯基','契诃夫','勃洛克','阿赫马托娃','阿赫玛托娃','茨维塔耶娃','曼德尔施塔姆','马雅可夫斯基','赫列勃尼科夫','帕斯捷尔纳克','布宁','高尔基','肖洛霍夫','布尔加科夫','帕乌斯托夫斯基','索尔仁尼琴','纳博科夫','布罗茨基']
    author_denmark = ['安徒生']
    author_austria = ['霍夫曼·斯塔尔','茨威格','卡夫卡','里尔克','特拉克尔','策兰']
    author_czech = ['聂鲁达','伏契克','米兰·昆德拉']
    author_germany = ['歌德','莱辛','席勒','海涅','尼采','布莱希特']
    author_hungary = ['裴多菲','伊姆莱因','凯尔泰斯','莫里兹']
    author_poland = ['显克维支','莱蒙特','辛波丝']
    author_switzerland = ['迪伦马特','马丁·苏塔']
    author_italy = ['卡尔维诺','科洛迪','莫拉维亚','邓南遮']
    author_romania = ['普列达','布楚拉']
    author_serbia = ['帕维奇']
    author_spain = ['塞万提斯','马丁内斯','埃切加赖','希梅内斯','莫里纳']
    author_egypt = ['贾巴拉','塔哈·侯赛因']
    author_australia = ['考琳·麦卡洛','帕特利克·怀特','克里·格林伍德']
    author_us = ['亨利·梭罗','马克·吐温','欧·亨利','霍桑','惠特曼','爱伦·坡','海明威','富兰克林','艾默生','爱默生','亨利·詹姆斯','狄金森','德莱塞','弗罗斯特','奥尼尔','福克纳']
    author_canada = ['罗伯特·索耶','蒙哥马利']
    author_columbia = ['马尔克斯']
    author_venezuela = ['皮耶德里']
    author_uruguay = ['加利亚诺']
    author_argentina = ['博尔赫斯','拉雷塔','玻塞']
    author_chile = ['聂鲁达']
    author_peru = ['尤萨']

    name_in_text = []

    for name in author_japan:
        if name in text:
            name_in_text.append('[日本]+'+name)

    for name in author_india:
        if name in text:
            name_in_text.append('[印度]+' + name)

    for name in author_belgium:
        if name in text:
            name_in_text.append('[比利时]+' + name)

    for name in author_france:
        if name in text:
            name_in_text.append('[法国]+' + name)

    for name in author_uk:
        if name in text:
            name_in_text.append('[英国]+' + name)

    for name in author_belarus:
        if name in text:
            name_in_text.append('[白俄罗斯]+' + name)

    for name in author_russia:
        if name in text:
            name_in_text.append('[俄罗斯]+' + name)

    for name in author_denmark:
        if name in text:
            name_in_text.append('[丹麦]+' + name)

    for name in author_austria:
        if name in text:
            name_in_text.append('[奥地利]+' + name)

    for name in author_czech:
        if name in text:
            name_in_text.append('[捷克]+' + name)

    for name in author_germany:
        if name in text:
            name_in_text.append('[德国]+' + name)

    for name in author_hungary:
        if name in text:
            name_in_text.append('[匈牙利]+' + name)

    for name in author_poland:
        if name in text:
            name_in_text.append('[波兰]+' + name)

    for name in author_switzerland:
        if name in text:
            name_in_text.append('[瑞士]+' + name)

    for name in author_italy:
        if name in text:
            name_in_text.append('[意大利]+' + name)

    for name in author_romania:
        if name in text:
            name_in_text.append('[罗马尼亚]+' + name)

    for name in author_serbia:
        if name in text:
            name_in_text.append('[塞尔维亚]+' + name)

    for name in author_spain:
        if name in text:
            name_in_text.append('[西班牙]+' + name)

    for name in author_egypt:
        if name in text:
            name_in_text.append('[埃及]+' + name)

    for name in author_australia:
        if name in text:
            name_in_text.append('[澳大利亚]+' + name)

    for name in author_us:
        if name in text:
            name_in_text.append('[美国]+' + name)

    for name in author_canada:
        if name in text:
            name_in_text.append('[加拿大]+' + name)

    for name in author_columbia:
        if name in text:
            name_in_text.append('[哥伦比亚]+' + name)

    for name in author_venezuela:
        if name in text:
            name_in_text.append('[委内瑞拉]+' + name)

    for name in author_uruguay:
        if name in text:
            name_in_text.append('[乌拉圭]+' + name)

    for name in author_argentina:
        if name in text:
            name_in_text.append('[阿根廷]+' + name)

    for name in author_chile:
        if name in text:
            name_in_text.append('[智利]+' + name)

    for name in author_peru:
        if name in text:
            name_in_text.append('[秘鲁]+' + name)

    return name_in_text

def excel_data(file):
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(file)
        # 获取第一个工作表
        table = data.sheet_by_index(0)
        # 获取行数
        nrows = table.nrows
        # 获取列数
        ncols = table.ncols

        new_workbook = copy(data)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)
        new_worksheet.write(0, 13, "被研究作家国别")
        new_worksheet.write(0, 14, "被研究作家名")
        for row in range(1, nrows):
            cell_value1 = table.cell(row, 0).value
            cell_value2 = table.cell(row, 6).value
            cell_value3 = table.cell(row, 7).value
            # 把数据追加到excel_list中
            info = cell_value1+'\n'+cell_value2+'\n'+cell_value3
            foreignAuthorList = getResearchedAuthors(info)
            nationality_list = []
            name_list = []
            for element in foreignAuthorList:
                nationality_list.append(element.split('+')[0])
                name_list.append(element.split('+')[-1])

            nationality = ';'.join(nationality_list)
            names = ';'.join(name_list)
            new_worksheet.write(row, 13, nationality)
            new_worksheet.write(row, 14, names)

        new_workbook.save(file)

excel_data('/Users/tyrzax/ZJU Documents/2524六大期刊-new.xls')



