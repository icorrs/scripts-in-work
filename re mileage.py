#coding = utf8
#use re to find calculate mileage in calculate sheet

import re

import pandas as pd
import openpyxl 

import get_path
import BOQ_list_to_formulate_1_2

pat = r'([ABCDZY]{0,1}K(\d+)\+(\d+))'
def find_mileage():
    path = get_path.getpath('calculate xlsx file')
    boq_dic = BOQ_list_to_formulate_1_2.get_boq_dic()
    work_book = openpyxl.load_workbook(path)
    long_word_list = []
    for sh in work_book: 
        sh['e5'] = boq_dic[sh['b4'].value][1]
        cal_word = sh['a10'].value
        if len(cal_word)>=800:
            long_word_list.append(sh.title)
        if re.search(pat,sh['g4'].value):
            mileage_list = re.findall(pat,sh['g4'].value)
            frame1 = pd.DataFrame(mileage_list,columns=['mileage','mile','submile'])
            frame1['order_col'] = frame1['mile'].apply(lambda x:int(x)*1000)+\
                frame1['submile'].apply(lambda x:float(x))
            frame1.sort_values(by='order_col')
            sh['b5'] = frame1['mileage'][0]
            sh['c5'] = frame1['mileage'][len(frame1['mileage'])-1]
        else:
            sh['b5'] = 'K26+200'
            sh['c5'] = 'K41+600'
    work_book.save(path)
    print('以下表格计算式过长：')
    for j in long_word_list:
        print(j)


if __name__ == '__main__':
    find_mileage()

