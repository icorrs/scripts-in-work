"""
import excel exported from microsoft access to mysql,and change the title to english.
#读取由microsoft access导出的excel数据表到mysql，同时转换标题为英文。
"""

import os
import re
import datetime

import pandas as pd 
import sqlalchemy
import pymysql
import openpyxl


def get_title_dic():
    'get title\'s name in dic which key is chinese and value is english'
    path=input('please enter title_dic path:')
    title_dic={}
    if not os.path.isfile(path):
        print('wrong path,please enter title_dic path again:')
        get_title_dic()
    else:
        frame1=pd.read_excel(path)
        for i in range(len(frame1['中文'])):
            title_dic.setdefault(frame1.iloc[i,1],frame1.iloc[i,2])
    return title_dic


def mysql_engine():
    'get host,port,user,password,database to mysql,return sqlalchemy engine'
    (user,password,host,port,database)=tuple((input('please enter (user,password,host,port,database):').split(',')))
    engine=sqlalchemy.create_engine('mysql+pymysql://%s:%s@%s:%d/%s?charset=utf8'%(user,password,host,int(port),database))
    if engine:
        return engine
    else:
        mysql_engine()

def get_path():
    'get path that contains the excel file need to be imported'
    path=input('please enter excel file path:')
    if os.path.isdir(path):
        return path
    else:
        get_path()


def mysql_mirage():
    'mirage main attribute'
    title_dic=get_title_dic()
    engine=mysql_engine()
    path=get_path()
    os.chdir(path)
    file_list=os.listdir(path)
    for file in file_list:
        file_name=os.path.splitext(file)[0]
        if os.path.splitext(file)[1][1:] in ('xlsx','xls'):
            frame_file=pd.read_excel(file)
            frame_file=frame_file.rename(title_dic,axis='columns')
            print('%s begin to mirage %s to mysql'%(str(datetime.datetime.today()),file))
            if file_name in title_dic.keys():
                frame_file.to_sql(title_dic[file_name].replace(' ',''),engine,if_exists='replace',index=False)
            else:
                frame_file.to_sql(file_name,engine,if_exists='replace',index=False)
        else:
            pass
    print('%s finished'%(str(datetime.datetime.today())))


if __name__=='__main__':
    mysql_mirage()
