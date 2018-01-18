#sub-contractor detail_wbs join income_wbs and sub_contract boq
import pandas as pd
frame_subcontract_wbs=pd.read_excel(r'c:\python tem\分包量连接wbs.xlsx')
frame_subcontract_statement=pd.read_excel(input('please enter sub_contract statement detail_wbs:'))
def sub_contract_detail_wbs_join(frame1):
    frame_wbs=pd.read_excel(r'c:\python tem\wbs.xlsx')
    frame_subcontract_boq=pd.read_excel(r'c:\python tem\sub-contract boq.xlsx')
    frame1=frame1.pivot_table(['分包清单量'],index=['WBS序号','分包清单编号'],columns=['分包商代号'],aggfunc='sum')['分包清单量'].reset_index()
    frame1=pd.merge(frame1,frame_wbs,on=['WBS序号'],how='outer')
    frame1=pd.merge(frame1,frame_subcontract_boq,on=['分包清单编号'],how='outer')
    return frame1
sub_contract_detail_wbs_join(frame_subcontract_wbs).to_csv(r'c:\python tem\sub_contract_wbs.csv',encoding='utf-8-sig')
sub_contract_detail_wbs_join(frame_subcontract_statement).to_csv(r'c:\python tem\sub_contract_statement_wbs.csv',encoding='utf-8-sig')
