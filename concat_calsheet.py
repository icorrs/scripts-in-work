#concate income calculate sheets exported from calculate system to 1 sheet

import pandas as pd 
import openpyxl

import get_path

def join_sheet():
    path = get_path.getpath('输入文件')
    path_out = get_path.getdir('输出文件')
    wb = openpyxl.load_workbook(filename=path)
    list_sheetsname = wb.get_sheet_names()
    list_title = ['清单编号','项目名称','中间计量表编号',
    '部位','单位','数量','单价（元）','金额']
    list_word = ['a','b','d','e','g','h','i','j']
    list_dic = []
    for sheet_name in list_sheetsname:
        sh = wb.get_sheet_by_name(sheet_name)
        for i in range(5,29):
            dic_row = {}
            for j in range(8):
                dic_row.setdefault(list_title[j],sh[list_word[j]+str(i)].value)
            list_dic.append(dic_row)
    frame_out = pd.DataFrame(list_dic)
    frame_out.to_csv(path_out+'\joinsheet.csv',encoding='utf-8-sig')


if __name__ == '__main__':
    join_sheet()
