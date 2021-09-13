# -*- coding: utf-8 -*-
import os, sys
import pdftotext
import xlrd
import xlwt
from xlutils.copy import copy
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import re
from tkinter import filedialog
# from xpinyin import Pinyin
import numpy as np
import jieba
import jieba.posseg as pseg
import csv
import codecs

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

def danwei(institution_txt):
    institution_txt_list = []
    institution_txt = institution_txt.replace(',', '，').replace('。', '，')
    if len(institution_txt.split('，')) > 1:
        institution_txt_list = institution_txt.split('，')
    else:
        institution_txt_list.append(institution_txt)
    institution_list = []
    institution_remove_list = []
    for word in institution_txt_list:
        if '研究所' in word:
            institution_list.append(word.split('研究所')[0] + '研究所')
        elif '系' in word and ('学院' in word or '大学' in word):
            institution_list.append(word.split('系')[0] + '系')
        elif '学院' in word:
            institution_list.append(word.split('学院')[0] + '学院')
        elif '大学' in word:
            institution_list.append(word.split('大学')[0] + '大学')
        if ('学报' in word or '出版社' in word or '科研' in word or '成果' in word) and word in institution_list:
            institution_list.remove(word)
    institution_list = list(set(institution_list))
    for word in institution_list:
        for key in institution_list:
            if (word != key) and (word in key):
                print(institution_list)
                print(word)
                print(key)
                institution_remove_list.append(word)
    institution_remove_list = list(set(institution_remove_list))
    print(institution_remove_list)
    for word in institution_remove_list:
        institution_list.remove(word)
    if institution_list == []:
        institution = '找不到【作者单位】！'
    else:
        institution = '；'.join(institution_list)
    return institution


def longest_substring(str1, str2):
    # 首先创建一个长宽分别为len(str1),len(str2)的二维数组
    record = np.zeros(shape=(len(str1), len(str2)))
    # 获取第一个字符串的长度
    str1_length = len(str1)
    # 获取第二个字符串的长度
    str2_length = len(str2)
    # 最大长度
    maxLen = 0
    # 结束的索引
    maxEnd = 0
    for i in range(str1_length):
        for j in range(str2_length):
            # 判断两个字符串对应的索引是否相同
            if str1[i] == str2[j]:
                # 判断是否是第一行或者是第一列
                if i == 0 or j == 0:
                    # 如果是则对应索引置一
                    record[i][j] = 1
                else:
                    # 如果不是对应的索引则为其左上角对应的元素加一
                    record[i][j] = record[i - 1][j - 1] + 1
            else:
                # 如果字符串对应的元素不相同则置零
                record[i][j] = 0
            # 判断记录数值是否大于最大长度
            if record[i][j] > maxLen:
                # 如果是则充值最大长度
                maxLen = record[i][j]
                # 将结束索引置为i
                maxEnd = i
    # print(record)
    # 返回结果为，起始的索引及最大的长度
    return maxEnd - maxLen + 1, maxEnd + 1


def cn_authortitle(filename, filetxt, author_title_txt):


    if '摘..要' in author_title_txt:
        at_txt = author_title_txt.split('摘..要')[0]

        author_txt = at_txt.strip('\n').strip().split('\n')[-1]
        if len(author_txt.replace(' ', '')) > 8:
            at_txt = at_txt.replace(author_txt, '')
            author_txt = at_txt.strip('\n').strip().split('\n')[-1]
        # author = author_txt
        # author = filename.split('_')[-1].replace('.txt','')
        title_txt = at_txt.replace(author_txt, '').replace('\n', '').replace(' ', '')
        title = title_txt
        # filename_title = filename.replace('_'+author+'.txt','')
        author = author_txt.replace(' ', '')
        if ' ' in author_txt:
            author_list = author_txt.split(' ')
            author_list = [i for i in author_list if (len(str(i)) != 0)]
            author = author_list[0]
            author2 = author_list[1]
        else:
            author2 = ''
        if len(author) < 2:
            author = author_txt.replace(' ', '')
            author2 = ''

        if len(author) < 2 or len(author) > 10:
            author = filename.split('_')[-1].replace('.txt', '')
            author2 = ''

    else:
        author = filename.split('_')[-1].replace('.txt', '')
        author2 = ''
        if '-' in filename:
            title = filename.split('-')[-1].replace('_' + author, '').replace('.txt', '')
        else:
            title = filename.replace('_' + author, '').replace('.txt', '')

    if '”' in title and '“' not in title:
        title = '“' + title
    if '*' in title:
        title = title.replace('*', '')
    if '———' in title:
        title = title.replace('———', '——')
    if '——·' in title:
        title = title.replace('——·', '·')
    if '—' in title and '——' not in title:
        title = title.replace('—', '——')
    if '⊙' in title:
        title = title.replace('⊙', '')
    if len(title) < 2 or len(title) > 100:
        if '-' in filename:
            title = filename.split('-')[-1].replace('_' + author, '').replace('.txt', '')
        else:
            title = filename.replace('_' + author, '').replace('.txt', '')

    # print(title)
    if 'EN..TITLE' in filetxt and 'EN..ABSTRACT' in filetxt:
        en_title = filetxt.split('EN..TITLE')[1].split('EN..ABSTRACT')[0].replace('\n', '')
    elif 'EN..TITLE' in filetxt:
        en_title = filetxt.split('EN..TITLE')[1].split('\n')[0]
    else:
        en_title = ''

    # author = filename.split('_')[-1].replace('.txt','')


    return author, author2, title, en_title


def country(file):
    with open(file, 'r') as f:
        txt = f.read()
    filename = file.split("/")[-1].split('.')[0]
    # print(filename)

    AsiaCountryList = ['中国', '日本', '韩国', '朝鲜', '蒙古', '越南', '老挝', '柬埔寨', '缅甸', '马来西亚', '新加坡', '泰国', '菲律宾', '印尼', '东帝汶',
                       '巴基斯坦', '印度', '不丹', '尼泊尔', '孟加拉国', '马尔代夫', '斯里兰卡', '伊朗', '伊拉克', '科威特', '约旦', '沙特阿拉伯', '卡塔尔',
                       '巴林', '阿联酋', '也门', '阿曼', '以色列', '阿塞拜疆', '格鲁吉亚', '亚美尼亚', '土耳其', '阿富汗', '黎巴嫩', '吉尔吉斯斯坦', '土库曼斯坦',
                       '塔吉克斯坦', '乌兹别克斯坦', '文莱']
    WesternEuropeList = ['比利时', '法国', '爱尔兰', '卢森堡', '摩纳哥', '荷兰', '英国']
    NortheastEuropeList = ['白俄罗斯', '爱沙尼亚', '拉脱维亚', '立陶宛', '摩尔多瓦', '俄罗斯', '乌克兰', '丹麦', '芬兰', '冰岛', '瑞典']
    CentralandSouthernEuropeList = ['奥地利', '捷克', '德国', '匈牙利', '列支敦士登', '波兰', '斯洛伐克', '瑞士', '阿尔巴尼亚', '保加利亚', '克罗地亚',
                                    '希腊', '意大利', '马其顿', '马耳他', '葡萄牙', '罗马尼亚', '塞尔维亚', '斯洛文尼亚', '西班牙']
    AfricaList = ['阿尔及利亚', '安哥拉', '贝宁', '博茨瓦纳', '布基纳法索', '布隆迪', '佛得角', '喀麦隆', '中非', '乍得', '科摩罗', '科特迪瓦', '刚果民主共和国',
                  '吉布提', '埃及', '赤道几内亚', '厄立特里亚', '埃塞俄比亚', '加蓬', '冈比亚', '加纳', '几内亚', '几内亚比绍', '肯尼亚', '莱索托', '利比里亚',
                  '利比亚', '马达加斯加', '马拉维', '马里', '毛里塔尼亚', '毛里求斯', '马约特', '莫桑比克', '摩洛哥', '纳米比亚', '尼日尔', '尼日利亚', '刚果共和国',
                  '卢旺达', '圣多美和普林西比', '塞内加尔', '塞舌尔', '塞拉利昂', '索马里', '南非', '南苏丹', '苏丹', '斯威士兰', '坦桑尼亚', '多哥', '突尼斯',
                  '乌干达', '赞比亚', '津巴布韦']
    OceaniaList = ['澳大利亚', '巴布亚新几内亚', '新西兰', '斐济', '所罗门群岛', '瓦努阿图', '萨摩亚', '基里巴斯', '密克罗尼西亚联邦', '汤加', '马绍尔群岛', '帕劳',
                   '图瓦卢', '瑙鲁']
    USList = ['美国']
    CanadaandotherAmericanCountriesList = ['加拿大', '墨西哥', '古巴', '海地', '多米尼加', '多米尼克', '格林纳达', '圣文森特和格林纳丁斯', '圣卢西亚',
                                           '圣基茨和尼维斯', '安提瓜和巴布达', '巴巴多斯', '伯利兹', '萨尔瓦多', '尼加拉瓜', '哥斯达黎加', '巴哈马', '巴拿马',
                                           '洪都拉斯', '哥伦比亚', '委内瑞拉', '玻利维亚', '秘鲁', '厄瓜多尔', '圭亚那', '苏里南', '巴西', '巴拉圭',
                                           '乌拉圭', '阿根廷', '智利', '危地马拉', '牙买加', '特立尼达和多巴哥']

    frequency_list = []
    # Asia
    AsiaCount = 0
    for AsiaCountry in AsiaCountryList:
        AsiaCount += len(re.findall(AsiaCountry, txt))
    # country_list.append('AsiaCountry:')
    # frequency_list.append(AsiaCount)
    frequency_list = {'亚洲': AsiaCount}

    # Western Europe
    WesternEuropeCount = 0
    for WesternEuropeCountry in WesternEuropeList:
        WesternEuropeCount += len(re.findall(WesternEuropeCountry, txt))
    # country_list.append('WesternEuropeCountry:')
    # frequency_list.append(WesternEuropeCount)
    frequency_list['西欧'] = WesternEuropeCount

    # Northeast Europe
    NortheastEuropeCount = 0
    for NortheastEuropeCountry in NortheastEuropeList:
        NortheastEuropeCount += len(re.findall(NortheastEuropeCountry, txt))
    # country_list.append('NortheastEuropeCountry:')
    # frequency_list.append(NortheastEuropeCount)
    frequency_list['东欧、北欧'] = NortheastEuropeCount

    # Central and Southern Europe
    CentralandSouthernEuropeCount = 0
    for CentralandSouthernEuropeCountry in CentralandSouthernEuropeList:
        CentralandSouthernEuropeCount += len(re.findall(CentralandSouthernEuropeCountry, txt))
    # country_list.append('CentralandSouthernEuropeCountry:')
    # frequency_list.append(CentralandSouthernEuropeCount)
    frequency_list['中欧、南欧'] = CentralandSouthernEuropeCount

    # Africa
    AfricaCount = 0
    for AfricaCountry in AfricaList:
        AfricaCount += len(re.findall(AfricaCountry, txt))
    # country_list.append('AfricaCountry:')
    # frequency_list.append(AfricaCount)
    frequency_list['非洲'] = AfricaCount

    # Oceania
    OceaniaCount = 0
    for OceaniaCountry in OceaniaList:
        OceaniaCount += len(re.findall(OceaniaCountry, txt))
    # country_list.append('OceaniaCountry:')
    # frequency_list.append(OceaniaCount)
    frequency_list['大洋洲'] = OceaniaCount

    # United States
    USCount = 0
    for USCountry in USList:
        USCount += len(re.findall(USCountry, txt))
    # country_list.append('USCountry:')
    # frequency_list.append(USCount)
    frequency_list['美国'] = USCount

    # Canada and other American Countries
    CanadaandotherAmericanCountriesCount = 0
    for CanadaandotherAmericanCountries in CanadaandotherAmericanCountriesList:
        CanadaandotherAmericanCountriesCount += len(re.findall(CanadaandotherAmericanCountries, txt))
    # country_list.append('CanadaandotherAmericanCountries:')
    # frequency_list.append(CanadaandotherAmericanCountriesCount)
    frequency_list['加拿大及其他美洲国家'] = CanadaandotherAmericanCountriesCount

    # count_list = ['亚洲:','西欧:','东欧、北欧:','中欧、南欧:','非洲:','大洋洲:','美国:','加拿大及其他美洲国家:',"未归类："]

    # print(frequency_list)
    max_frequency = max(frequency_list, key=frequency_list.get)
    # print(max_frequency)

    return max_frequency, frequency_list


def write_excel_xls_append(path, value):
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets[0])
    rows_old = worksheet.nrows
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(0)
    new_worksheet.write(rows_old, 0, value)
    new_workbook.save(path)

def cutspace(ifn, ofn):
    infile = open(ifn, 'r')
    outfile = open(ofn, 'w')

    content = infile.read()
    content = content.replace('\n\n', '')
    content = content.replace('  ', '')
    outfile.write(content)

    infile.close
    outfile.close
    if (os.path.exists(ifn)):
        os.remove(ifn)
        # print '移除后test 目录下有文件：%s' %os.listdir(dirPath)
    else:
        print("要删除的文件不存在！")


def pdf2txt(pdfpath, txtpath):
    path = '/Users/tyrzax/ZJU Documents/error_open.xls'
    with open(pdfpath, "rb") as f:
        try:
            pdf = pdftotext.PDF(f)
        except Exception:
            write_excel_xls_append(path, pdfpath + '无法打开')
            return

    with open(txtpath, 'w') as f:
        f.write("\n-----next page-----\n".join(pdf))




def get_filelist(path, type):
    Filelist = []
    for home, dirs, files in os.walk(path):
        for filename in files:
            if filename.split('.')[-1] == type:
                Filelist.append(os.path.join(home, filename))
    return Filelist


def output(pdfpath):
    Filelist_pdf = get_filelist(pdfpath, 'pdf')
    Filelist_txt = get_filelist(pdfpath, 'txt')

    for filename in Filelist_pdf:
        if filename.replace('.pdf','.txt') in Filelist_txt:
            print('已存在!\n' + filename)
        else:
            print('正在处理:' + filename)
            print('还剩：' + str(len(Filelist_pdf) - len(get_filelist(pdfpath, 'txt')) - 1))
            pdf2txt(filename, filename.replace('.pdf','.txt'))
    print(len(Filelist_pdf))

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
    name = text.split('.')[0]
    with codecs.open(name+'.csv','w')as f:
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


def lunwen(filepath, m):
    filename = filepath.split("/")[-1]
    print(filename)
    with open(filepath, 'r') as f:
        filetxt = f.read()

    filetxt = filetxt.replace('＊', '*')
    zhaiyao_list = ['内 容提 要', '摘      要 ：', '摘    要 ：', '[摘     要 ]', '［摘        要］', '［摘       要］', '［摘      要］',
                    '［摘     要］', '［摘    要］', '［内容提要］', '［内容摘要］', '［提 要］', '［提要］', '［摘 要］', '［摘要］', '【内容提要】', '【内容摘要】',
                    '【提 要】', '【提要】', '【摘 要】',
                    '摘       要：', '【摘要】', '内容提要：', '内容提要:', '内容提要|', '内容摘要：', '内容摘要:', '内容摘要|', '提 要：', '提 要:', '提 要|',
                    '提要：', '提要:',
                    '提    要：', '提   要：', '提  要：', '提要|', '摘 要：', '摘 要:', '摘 要|', '摘要：', '摘要:', '摘要|', '内容提要', '内容摘要',
                    '〔内容提要〕', '〔内容摘要〕', '〔提 要〕', '〔提要〕', '〔摘 要〕', '〔摘要〕', '摘   要：', '摘   要:', '摘   要', '提   要 |',
                    '【提   要】', '提   要', '【提     要】', '提     要', '摘    要', '摘       要', '摘  要', '摘     要', '要］',
                    '摘      要：', '摘     要：', '摘    要：', '摘  要：', '提 要', '提要', '摘 要', '摘要']
    guanjianci_list = ['关 键词', '关 键 词', '【关 键 词】', '[ 关键词 ]', '关键词：', '关键词:', '［关键词］', '关键词|', '【关键词】', '〔关键词〕',
                       '关键词 |', '关键词 ：', '关键词 :', '关   键   词：', '键词］', '主题词', '关键词']
    danwei_list = ['作 者 简 介：', '[ 作者简介 ]', '[作者单位]', '[作者简介]', '[作者信息]', '作者单位：', '作者单位:', '作者单位|', '作者单位 ：', '作者单位 :',
                   '作者单位 |', '作者简介：',
                   '作者简介:', '作者简介|', '作者简介 ：', '作者简介 :', '作者简介 |', '作者信息：', '作者信息:', '作者信息|', '作者信息 ：', '作者信息 :',
                   '作者信息 |', '【作者单位】', '【作者简介】', '【作者信息】', '〔作者单位〕', '〔作者简介〕', '〔作者信息〕', '［作者简介］', '［作者单位］', '［作者信息］',
                   '作者简介：', '本 文 作 者 ：', '本文作者：', '作者单位', '作者简介']
    entitle_list = ['Title:', 'TITLE:', 'Title：', 'TITLE：']
    enzhaiyao_list = ['Abstract：', 'Abstract:', 'ABSTRACT：', 'ABSTRACT:']
    biaoti_list = ['标题：', '标题:', '题目：', '题目:', '标题 :', '题目 :', '题目 ：']
    enzuozhe_list = ['Authors: ', 'Author:']

    qikan_list = ['No.', 'NO.', 'DOI', '2018', '二○一八', '期', '辑', '第', '二零一八']


    for word in zhaiyao_list:
        filetxt = filetxt.replace(word, '摘..要')
    for word in guanjianci_list:
        filetxt = filetxt.replace(word, "关..键..词")
    for word in danwei_list:
        filetxt = filetxt.replace(word, "单..位")
    for word in entitle_list:
        filetxt = filetxt.replace(word, "EN..TITLE")
    for word in biaoti_list:
        filetxt = filetxt.replace(word, "标..题")
    for word in enzuozhe_list:
        filetxt = filetxt.replace(word, "EN..AUTHOR")
    for word in enzhaiyao_list:
        filetxt = filetxt.replace(word, 'EN..ABSTRACT')

    filetxt = filetxt.replace('(', '（')
    filetxt = filetxt.replace(')', '）')
    first_page = filetxt.split('-----next page-----')[0]

    # # 第一页和第二页
    page12 = filetxt
    if '-----next page-----' in filetxt:
        page12 = 'next page tag'.join(filetxt.split('-----next page-----')[0:2]).replace('\n', '（换行标记）')
        nextpage_txt = re.findall('next page tag.*?（换行标记）（换行标记）', page12)[0]
        page12 = page12.replace(nextpage_txt, '').replace('（换行标记）', '\n')

    author_title_txt = page12.replace(' ', '')
    for word in zhaiyao_list:
        author_title_txt = author_title_txt.replace(word, '摘..要')
    author, author2, title, en_title = cn_authortitle(filename, filetxt, author_title_txt)


    # 3. abstract
    abstract_txt = page12.replace(' ', '')
    for word in zhaiyao_list:
        abstract_txt = abstract_txt.replace(word, '摘..要')
    if '摘..要' in abstract_txt and '关..键..词' in abstract_txt:
        abstract_txt = abstract_txt.split('摘..要')[1].split('关..键..词')[0]
        abstract = abstract_txt.replace('\n', '').replace(' ', '').replace('［', '').replace('］', '')
    elif '摘..要' in abstract_txt and '\n\n\n\n' in abstract_txt:
        abstract_txt = abstract_txt.split('摘..要')[1].split('\n\n\n\n')[0]
        abstract = abstract_txt.replace('\n', '').replace(' ', '').replace('［', '').replace('］', '')
    else:
        abstract = ''
    if abstract[0] == ':':
        abstract = abstract.replace(abstract[0], '')

    # print(abstract)

    # 4. keywords
    # keywords_txt = page12.replace(' ','')
    keywords_txt = page12
    if '关..键..词' in keywords_txt and 'DOI' in keywords_txt:
        keywords_txt = keywords_txt.split('关..键..词')[1].split('DOI')[0].replace('\n', '')
        if '；' in keywords_txt:
            keywords = '；'.join(keywords_txt.replace('\n', '').replace(' ', '').split('；'))
        elif '，' in keywords_txt:
            keywords = '；'.join(keywords_txt.replace('\n', '').replace(' ', '').split('，'))
        else:
            keywords = '；'.join(keywords_txt.replace('\n', '').strip(' ').split())
        keywords = keywords.replace('，', '；')
    elif '关..键..词' in keywords_txt and '\n' in keywords_txt.split('关..键..词')[1]:
        keywords_txt = keywords_txt.split('关..键..词')[1].split('\n')[0].replace('\n', '')
        if '；' in keywords_txt:
            keywords = '；'.join(keywords_txt.replace('\n', '').replace(' ', '').split('；'))
        elif '，' in keywords_txt:
            keywords = '；'.join(keywords_txt.replace('\n', '').replace(' ', '').split('，'))
        else:
            keywords = '；'.join(keywords_txt.replace('\n', '').strip(' ').split())
        keywords = keywords.replace('，', '；')
    else:
        keywords = ''
    keywords.replace('；；', '；')

    # 5. institution
    institution_txt = filetxt.replace(' ', '')
    if '单..位' in institution_txt:
        institution_txt = filetxt.replace(' ', '').split('单..位')[1].split('\n\n')[0].replace('\n', '')
    else:
        institution_txt = ''.join(institution_txt.split('-----next page-----')[-2:]).replace('\n', '')
    if author2 != '':
        if '英文' in filename:
            if '。' in institution_txt:
                institution_txt1 = institution_txt.split('。')[0]
                institution_txt2 = institution_txt.split('。')[1]
            elif '；' in institution_txt:
                institution_txt1 = institution_txt.split('；')[0]
                institution_txt2 = institution_txt.split('；')[1]
            else:
                institution_txt1 = institution_txt
                institution_txt2 = ''
            institution = danwei(institution_txt1)
            institution2 = danwei(institution_txt2)
        else:
            if author2 in institution_txt:
                institution_txt1 = institution_txt.split(author2)[0]
                institution_txt2 = institution_txt.split(author2)[1]
            else:
                institution_txt1 = institution_txt
                institution_txt2 = ''
            institution = danwei(institution_txt1)
            institution2 = danwei(institution_txt2)
    else:
        institution = danwei(institution_txt)
        institution2 = ''
    if institution2 != '' and institution2[0] == ',':
        institution2 = institution2.replace(institution2[0], '')

    # 6. journal
    journal_txt = filetxt.split('-----next page-----')[0:3]
    journal_txt = ''.join(journal_txt).replace(' ', '').split('\n')
    filename_journal = '当代外国文学'
    journal_txt_list = []
    journal_list = []
    for word in journal_txt:
        if filename_journal in word:
            journal_txt_list.append(word)
        if 'DOI' in word:
            journal_txt_list.append(word)
    # j name
    journal_list.append('《' + filename_journal + '》')
    # print(journal_txt)


    journal_year = filepath.split('/')[-2].split('-')[0]
    journal_date = filepath.split('/')[-2].split('-')[1]

    journal_list.append(journal_year + '年第' + journal_date + '期')

    journal = '，'.join(journal_list)
    # print(journal)

    # 7. project
    # project_txt = page12.replace(' ', '')
    project_txt = filetxt.replace(' ', '')
    xiangmu_list = ['*本文为', '*本文系', '*本文是', '本文系', '本文为', '本文是', '本论文为', '本论文系', '本论文是', '本文得到', '本文已得到']
    # xiangmu_list = ['基金项目：','基金项目:','项目基金：','项目基金:']
    project = ''



    # 中外文学研究
    if re.findall("\*本文系.*?成果。", project_txt) != []:
        # print(re.findall("\*本文系.*?成果。", project_txt))
        project = project + re.findall("\*本文系.*?成果。", project_txt)[0]
    if re.findall('\*本文系.*?资助。', project_txt) != []:
        project = project + re.findall("\*本文系.*?资助。", project_txt)[0]

    if project != '' and project[0] == ':':
        project = project.replace(project[0], '')
    if project != '' and project[0] == '］':
        project = project.replace(project[0], '')
    for word in project:
        if word in '（）－：／０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ．，“”':
            project = project.replace(word, '')
    project = project.replace('*', '').replace(' ', '')

    # 8. country
    country_max, frequency_list = country(filepath)
    country_list = str(frequency_list)

    f.close()


    infotext = title + '\n' + keywords + '\n' + abstract
    foreignAuthorList = getResearchedAuthors(infotext)
    nationality_list = []
    name_list = []
    for element in foreignAuthorList:
        nationality_list.append(element.split('+')[0])
        name_list.append(element.split('+')[-1])

    nationality = ';'.join(nationality_list)
    names = ';'.join(name_list)

    outputTable.write(m, 0, title)
    outputTable.write(m, 1, en_title)
    outputTable.write(m, 2, author)
    outputTable.write(m, 3, institution)
    outputTable.write(m, 4, author2)
    outputTable.write(m, 5, institution2)
    outputTable.write(m, 6, keywords)
    outputTable.write(m, 7, abstract)
    outputTable.write(m, 8, project)
    outputTable.write(m, 9, journal)
    outputTable.write(m, 10, country_max)
    outputTable.write(m, 11, country_list)
    outputTable.write(m, 12, nationality)
    outputTable.write(m, 13, names)




# Find files that ends with .txt
def find(obj):
    if obj.endswith(".txt"):
        path_list.append(obj)


def get_list_dir(dir_path):
    dir_files = os.listdir(dir_path)
    for file in dir_files:
        file_path = os.path.join(dir_path, file)
        if os.path.isfile(file_path):
            find(file_path)
        if os.path.isdir(file_path):
            get_list_dir(file_path)


# Extract text first
pdfpath = filedialog.askdirectory(initialdir=sys.path[0],
                                  title="请选择你想要提取的文件夹")
output(pdfpath)

# Get those txt file dir
path_list = []
get_list_dir(pdfpath)

# read error
readErrorfile = xlwt.Workbook(encoding='utf-8')
errtable = readErrorfile.add_sheet('data', cell_overwrite_ok=True)
errtable.write(0, 0, "ioerror title")
errtable.write(0, 1, "ioerror journal")
errtable.write(0, 2, "ValueError title")
errtable.write(0, 3, "ValueError journal")
errtable.write(0, 4, "IndexError title")
errtable.write(0, 5, "IndexError journal")
errtable.write(0, 6, "会议信息及投稿指南")

# Then find info
outputFile = xlwt.Workbook(encoding='utf-8')
outputTable = outputFile.add_sheet('data', cell_overwrite_ok=True)
outputTable.write(0, 0, "中文题名")
outputTable.write(0, 1, "英文题名")
outputTable.write(0, 2, "第一作者（含职称职务等信息）")
outputTable.write(0, 3, "机构")
outputTable.write(0, 4, "第二作者（含职称职务等信息）")
outputTable.write(0, 5, "机构")
outputTable.write(0, 6, "关键词")
outputTable.write(0, 7, "摘要")
outputTable.write(0, 8, "项目")
outputTable.write(0, 9, "期刊")
outputTable.write(0, 10, "国别")
outputTable.write(0, 11, "国别词频参考")
outputTable.write(0, 12, "被研究作家国别")
outputTable.write(0, 13, "被研究作家名")

i = 1
j = 1
k = 1
m = 1
for path in path_list:
    try:
        filename = path.split("/")[-1]
        filename_journal = path.split('/')[-2]
        if '会议信息' in filename or '投稿指南' in filename:
            errtable.write(i, 6, filename_journal + '/' + filename)
            i = i + 1
        else:
            lunwen(path, m)
            m = m + 1
            print("")
    except IOError as e:
        errtable.write(i, 0, filename)
        errtable.write(i, 1, filename_journal)
        i = i + 1

    except IndexError as e:
        errtable.write(k, 4, filename)
        errtable.write(k, 5, filename_journal)
        k = k + 1

text_path_list = []
with open('/Users/tyrzax/ZJU Documents/txt_list.txt','r') as txt_list:
    for line in txt_list:
        text_path_list.append(line)

for text_file in text_path_list:
    yanjiuwenti(text_file)


outputFile.save(pdfpath+'output.xls')
readErrorfile.save(pdfpath+'error.xls')