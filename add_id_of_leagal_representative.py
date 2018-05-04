#add id of leagal_representative in cost_analysis.sub_contractor_account of database
import openpyxl
import pandas as pd 
import sqlalchemy

import cost_analysis.cost_analysis as ca

engine_default = ca.engine_default
path_default = '''C:\python tem\分包商法定代表人身份证.xlsx'''


def add_id(path=path_default,engine=engine_default):
    frame1 = pd.read_excel(path)
    frame2 = pd.read_sql('select legal_representative from sub_contractor_account',engine_default)
    dic1 = {}
    for i in range(len(frame1['法定代表人'])):
        dic1.setdefault(frame1.iloc[i,1],frame1.iloc[i,2])
    print(dic1)
    for key in dic1:
        engine.execute('update sub_contractor_account set id_of_legal_representative=\'%s\' where legal_representative=\'%s\''%(dic1[key],key))


if __name__=='__main__':
    add_id()
