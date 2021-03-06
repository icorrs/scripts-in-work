#read cost_analysis excel data and write report that can be used in cost_analysis_report.

import os
import re

import pandas as pd
import numpy as np

def get_path():
    'get cost_analysis excel file path'
    path=input('please enter data_source route: ')
    if os.path.isfile(path):
        return path
    else:
        get_path()


def get_report():
    'get report'
    path=get_path()
    path_dir=os.path.dirname(path)
    file_name=os.path.basename(path)
    path_date=re.match(r'(\d+)\s(.+)',file_name).group(1)
    os.chdir(path_dir)
    frame_data=pd.read_excel(path)
    frame_data=frame_data.loc[:,['自编号','数值']]
    frame_data.loc[:,'数值']=frame_data.loc[:,'数值'].apply(lambda x:round(x,4))
    data_dic={}
    for i in range(1,46):  #存储成本分析报告所需数值为dict
        data_dic.setdefault(frame_data.iloc[i-1,0],frame_data.iloc[i-1,1])
    print('writting txt')
    f=open('%s_cost_analysis.txt'%(path_date),'w',encoding='utf8')
    f.write('''2.1:本期完成产值%s万元，占本期计划的%s%%；
              本年完成产值%s万元，占年度计划的%s%%；
              开累完成产值%s万元，占合同额%s万元的%s%%\n'''\
            %(data_dic['2.1.1'],data_dic['2.1.2']*100,\
              data_dic['2.1.3'],data_dic['2.1.4']*100,\
              data_dic['2.1.5'],data_dic['2.1.6'],data_dic['2.1.7']*100))
    f.write('''3.1:本期开累完成利润%s万元，利润率%s%%，预期项目总体利润%s万元，利润率%s%%，
               本期开累完成利润占总利润的%s%%。其中，不含100章本期开累完成利润%s万元，利润率%s%%，预期项目总体利润%s万元，利润率%s%%，
               本期开累完成利润占总利润的%s%%。其中，本期实际利润与预控成本利润对比差  /  万元，与下达的标后预算利润对比差  /  万元。
               （标后预算未下达，100章收入分劈未确定，以上利润数据不含100章）\n'''\
             %(data_dic['3.1.1'],data_dic['3.1.2']*100,data_dic['3.1.3'],data_dic['3.1.4']*100,\
               data_dic['3.1.5']*100,data_dic['3.1.6'],data_dic['3.1.7']*100,data_dic['3.1.8'],data_dic['3.1.9']*100,data_dic['3.1.10']*100))
    f.write('''4.1:本期已完工程实际收入为%s万元（不含100章累计实际收入为%s万元），
               其中：本期合同收入%s万元，工程变更收入%s万元，已完工未结算收入%s万元。占合同金额的%s%%。
               原因为部分已完工工程按计量规范和资料要求不能计量\n'''\
               %(data_dic['4.1.1'],data_dic['4.1.2'],\
                 data_dic['4.1.3'],data_dic['4.1.4'],data_dic['4.1.5'],data_dic['4.1.6']*100))
    f.write('''4.2:工程变更情况说明，累计已上报工程变更、索赔%s万元，已批复工程变更、索赔%s万元。
               特别对合同约定的材料调差方式为  总价包干  。累计已发生材料调差%s万元，预计后期还将发生材料调差%s万元，
               预计项目材料调差总额%s万元。未批复变更、索赔原因分析及应对措施\n'''\
             %(data_dic['4.2.1'],data_dic['4.2.2'],data_dic['4.2.3'],data_dic['4.2.4'],data_dic['4.2.5']))
    f.write('''5.1:直接费对比情况：本期标后预算直接费用   /  万元，本期实际成本直接费%s万元。
               对直接成本人工、船机、材料差值大的部分进行原因分析及应对措施\n'''\
             %(data_dic['5.1.1']))
    f.write('''5.4:船机费：本期自有船机设备资产折旧%s万元，燃油费/万元，占总额的  /  %%；
               本期外租设备%s万元，占总额的%s%%说明固定资产折旧情况和外租设备的使用情况说明，
               机械费盈亏对比分析，对燃油消耗过多的船机设备与定额进行分析，查找原因制定措施。\n'''\
             %(data_dic['5.4.1'],data_dic['5.4.2'],data_dic['5.4.3']*100))
    f.write('''6.1:安全文明施工费用情况说明：本期安全文明施工费实际发生%s万元，批复%s万元，占总额的%s%%。
              安全文明施工费情况分析及应对措施。\n'''\
             %(data_dic['6.1.1'],data_dic['6.1.2'],data_dic['6.1.3']*100))
    f.write('''6.2:大小临时工程情况说明：本期大临工程实际发生/万元，计量/万元，占总额的/%%，
                本期应摊销%s万元。本期小临工程实际发生/万元，计量/万元，占总额的  /  %%，
                本期应摊销%s万元。大小临工程费盈亏分析及应对措施。\n'''\
             %(data_dic['6.2.1'],data_dic['6.2.2']))
    f.write('''7.2:本期实际发生间接费%s万元，其中不可控间接费%s万元，可控间接费%s万元，占批复总额的  /  %%，
               占成本总额的%s%%。3.间接费偏高原因分析及应对措施。\n'''\
               %(data_dic['7.2.1'],data_dic['7.2.2'],data_dic['7.2.3'],data_dic['7.2.4']*100))
    f.write('''8.1:各项目部累计签订合同%s份，已上报总部审批合同共%s份，合同上报审批率%s%%，未上报总部审批合同共%s份。
               分析原因。未上报合同正在进行分包计划审批。\n'''\
             %(data_dic['8.1.1'],data_dic['8.1.2'],data_dic['8.1.3']*100,data_dic['8.1.4']))
    f.close()

if __name__=='__main__':
    get_report()
