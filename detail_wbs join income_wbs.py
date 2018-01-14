import pandas as pd
detail_wbs_route=input('please enter detail_wbs route:')
income_wbs_route=input('please enter income_wbs route:')
frame_detailwbs=pd.read_excel(detail_wbs_route)
frame_incomewbs=pd.read_excel(income_wbs_route)
frame_out=pd.merge(frame_detailwbs,frame_incomewbs,left_on='wbs序号',right_on='WBS序号',how='outer')
frame_out.to_csv(r'c:\python tem\detail_wbs join income_wbs.csv',encoding='utf-8-sig')