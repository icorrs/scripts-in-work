#2016-201803 临时工工资月明细表汇总出按人员台帐；
import pandas as pd 
import re
import os
def get_path():
    path_source=input("""
    please enter path:""")
    if os.path.isfile(path_source):
        return path_source
    else:
        print('wrong path,try again:')
        get_path()
def wage_sum():
    path=get_path()
    frame1=pd.read_excel(path,encoding='utf-8')
    frame1=frame1.loc[:,['姓名','职务','合计（元）','备注','月份']]
    frame1['姓名']=frame1['姓名'].map(lambda x:x.replace(' ','').strip().replace('（广西）','').replace('(广西)','').replace('韩帮旗','韩邦旗'))
    frame1['职务']=frame1['职务'].map(lambda x:str(x).replace('门岗','门卫'))
    len_frame=len(frame1['姓名'])
    for i in range(len_frame):
        if frame1.loc[i,'姓名']=='刘尊':
            frame1.loc[i,'职务']='施工员'
        elif frame1.loc[i,'姓名']=='费如兵':
            frame1.loc[i,'职务']='调度'
        else:
            pass
    frame1['合计（元）']=frame1['合计（元）'].map(lambda x:round(float(x),2))
    grouped=frame1.groupby(['姓名','职务'])
    frame2=grouped.agg({'合计（元）':['sum','mean','min','max','median'],'月份':['min','max','count']})
    frame2.to_csv(os.path.join(os.path.dirname(path),'earning.csv'),encoding='utf-8-sig')
    frame3=frame1.pivot_table(index=['月份'],values=['合计（元）'],aggfunc='sum')
    frame3.to_csv(os.path.join(os.path.dirname(path),'month_sum.csv'),encoding='utf-8-sig')
if __name__=='__main__':
    wage_sum()
