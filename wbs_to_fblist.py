<<<<<<< HEAD
import pandas as pd
import numpy as np
frame_wbs=pd.read_excel(r'C:\资料\贵隆项目\04 预算 成本\成本分析\201711 成本分析\20171120工程量基础.xlsx')
frame_wbs=frame_wbs.iloc[:,:20]
frame_out=frame_wbs.pivot_table(['分包清单量'],index=['总表序号','零号台账内容','零号台账起始桩号','零号台帐终止桩号','收入清单编号','收入清单名称','收入清单单位','收入清单量','分包清单编号','分包清单名称','分包清单单位'],columns=['分包商代号'],aggfunc=np.sum)
=======
import pandas as pd
import numpy as np
frame_wbs=pd.read_excel(r'C:\资料\贵隆项目\04 预算 成本\成本分析\201711 成本分析\20171120工程量基础.xlsx')
frame_wbs=frame_wbs.iloc[:,:20]
frame_out=frame_wbs.pivot_table(['分包清单量'],index=['总表序号','零号台账内容','零号台账起始桩号','零号台帐终止桩号','收入清单编号','收入清单名称','收入清单单位','收入清单量','分包清单编号','分包清单名称','分包清单单位'],columns=['分包商代号'],aggfunc=np.sum)
>>>>>>> fee347fd13a570677259dcba9a53ef747de71a1c
frame_out.to_csv(r'c:\python tem\a02 wbs to fblist.csv',encoding='utf-8-sig')