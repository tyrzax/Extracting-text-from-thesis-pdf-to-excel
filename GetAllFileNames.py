import os
import xlwt



filePath = '/Users/tyrzax/ZJU Documents/无法读取'
excelPath = '/Users/tyrzax/ZJU Documents/无法读取文件列表.xls'

list=[]
for i,j,k in os.walk(filePath):
    for element in k:
        n = '.DS_Store'
        m = '.pdf'
        if n or m in element:
            k.remove(element)
    for element in k:
        list.append(element)


def data_write(file_path, datas):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    i = 0
    for data in datas:
        sheet1.write(i,0,data)
        i = i + 1

    f.save(file_path)

data_write(excelPath,list)






