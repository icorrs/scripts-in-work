#read value of first line  of all the excel file in specified path and save as csv.\
#读取指定文件路径下excel文件第一行内容并保存为csv文件。

import os

import openpyxl
import pandas as pd 

# get the path
def get_path():
    path=input('please enter the path:')
    if os.path.isdir(path):
        return path
    else:
        get_path()


# main attribute
def title_coll():
    path=get_path()
    list_file=os.listdir(path)
    list_out=[]
    os.chdir(path)
    for file in list_file:
        if os.path.splitext(file)[1][1:] not in ('xlsx','xls'):
            pass
        else:
            frame1=pd.read_excel(file)
            list_out.append([os.path.splitext(file)[0]]+list(frame1.columns))
    list_out_set=[]
    for item in list_out:
        list_out_set.extend(item)
    list_out_set=list(set(list_out_set))
    list_out.append(list_out_set)
    frame_out=pd.DataFrame(list_out)
    frame_out.to_csv('%s/%s.csv'%(path,'title_collect'),encoding='utf-8-sig')


if __name__=='__main__':
    title_coll()
                 