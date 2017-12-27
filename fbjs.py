fbsdh=str(input('please enter fbsdh: '))#分包商代号
lhtz=str(input('please enter lhtz: '))#零号台账
import numpy as np
import pandas as pd 
frame_source=pd.read_excel(lhtz,encoding='utf-8')#读取零号台账为dataframe
frame_fb=frame_source[frame_source.loc[:,'分包商代号']==fbsdh].fillna(value=0)#索引拟分析分包商数据
frame_fb.loc[:,'本期分包量']=frame_fb.loc[:,'本期分包量'].apply(lambda x:round(x,2))
frame_hd=frame_fb[frame_fb.loc[:,'分包清单编号'].str.contains('421-(.+)')]#索引涵洞通道行
frame_hd=frame_hd.loc[:,['零号台账起始桩号', '分包清单编号', '本期分包量','分包清单名称']]#索引涵洞通道列
frame_hd_sv=frame_hd.pivot_table(['本期分包量'],index=['分包清单编号','分包清单名称'],columns=['零号台账起始桩号']).dropna(how='all',axis=1)#pivot
frame_hd_sv.to_csv(r'c:\python tem\%s_涵洞.csv'%(fbsdh),encoding='utf-8-sig')
frame_zjdz=frame_fb[frame_fb.loc[:,'零号台账内容'].str.contains('(桩基|墩柱)')].loc[:,['零号台账内容','零号台账起始桩号','零号台帐终止桩号','分包清单名称','分包清单编号','本期分包量']]
frame_zjdz_sv=frame_zjdz.pivot_table(['本期分包量'],index=['零号台账起始桩号','零号台帐终止桩号'],columns=['零号台账内容','分包清单编号','分包清单名称'])
frame_zjdz_sv.to_csv(r'c:\python tem\%s_桩基墩柱.csv'%(fbsdh),encoding='utf-8-sig')
frame_yzl=frame_source[frame_source.loc[:,'零号台账内容'].str.match(r'^预制梁')].loc[:,['零号台账内容','零号台账起始桩号','零号台帐终止桩号','分包清单名称','分包清单编号','本期分包量']]
frame_yzl=frame_yzl.pivot_table(['本期分包量'],index=['零号台账起始桩号','零号台帐终止桩号'],columns=['零号台账内容','分包清单编号','分包清单名称'])
frame_yzl.to_csv(r'c:\python tem\%s_预制梁.csv'%(fbsdh),encoding='utf-8-sig')
frame_other=frame_fb[frame_fb.loc[:,'分包清单编号'].str.contains('421-(.+)')!=1]#drop涵洞通道行
frame_other=frame_other[frame_other.loc[:,'零号台账内容'].str.contains('(桩基|墩柱)')!=1]#drop桩基墩柱行
frame_other=frame_other[frame_other.loc[:,'零号台账内容'].str.match(r'^预制梁')!=1]#drop预制梁行
frame_other=frame_other.loc[:,['分包清单编号','分包清单名称','零号台账内容','零号台账起始桩号','零号台帐终止桩号','本期分包量']]
frame_other=frame_other.sort_index(by=['分包清单编号','零号台账内容'])
frame_other=frame_other[frame_other.loc[:,'本期分包量']>0]
#frame_other.to_csv(r'c:\python tem\%s_其他汇总.csv'%(fbsdh),encoding='utf-8-sig') #其他项汇总，测试用；
dic_qdbh={}
qdbh_set=set(frame_other.iloc[:,0])
for qdbh in qdbh_set:
    frame_qdbh=frame_other[frame_other.loc[:,'分包清单编号']==qdbh]
    qdnr_set=set(frame_qdbh.loc[:,'零号台账内容'])
    cal_str=qdbh+':'
    for qdnr in qdnr_set:
        frame_qdnr=frame_qdbh[frame_qdbh.loc[:,'零号台账内容']==qdnr][['零号台账起始桩号','零号台帐终止桩号','本期分包量']]
        cal_str=cal_str+qdnr+':'
        for i in range(len(frame_qdnr.index)):
            cal_str=cal_str+str(frame_qdnr.iloc[i,2])+'('+str(frame_qdnr.iloc[i,0])+'~'+str(frame_qdnr.iloc[i,1])+')'+'+'
        cal_str=''.join(list(cal_str)[:-1])+'='+str(frame_qdnr.loc[:,'本期分包量'].sum())+';'
    cal_str=cal_str+'合计：'+str(frame_qdbh.loc[:,'本期分包量'].sum())
    dic_qdbh.setdefault(qdbh,cal_str)
frame_cal_str=pd.DataFrame(pd.Series(dic_qdbh))
#frame_cal_str.to_csv(r'c:\python tem\%s_其他计算式.csv'%(fbsdh),encoding='utf-8-sig') #计算式导出，测试用；
frame_other=frame_other.pivot_table(['本期分包量'],index=['分包清单编号','分包清单名称'],aggfunc=np.sum)
frame_cal_str=frame_cal_str.reindex(index=frame_other.index,level=0)
#frame_other.to_csv(r'c:\python tem\%s_other_pivot.csv'%(fbsdh),encoding='utf-8-sig') #其他项pivot后导出，测试用。
frame_other_sv=pd.merge(frame_other,frame_cal_str,left_index=True,right_index=True,how='outer')
frame_other_sv.to_csv(r'c:\python tem\%s_其他.csv'%(fbsdh),encoding='utf-8-sig')
