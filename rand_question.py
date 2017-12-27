import pandas as pd
import numpy as np
import openpyxl
i=int(input('please enter number of kinds the test should be:'))
j=int(input('please enter how many questions one test have:'))
f=input('please enter the question file route:',)
if not f:
    f=r'c:\python tem\formula.xlsx'
    print('no file route entered set file route as:%s'%(f))
wb=openpyxl.load_workbook(filename=f)
sh_list=wb.get_sheet_names()
sh1=wb.get_sheet_by_name(str(sh_list[0]))
question_dic={}
for k in range(1,j+1):
    question_dic.setdefault(sh1['a'+str(k+1)].value,[sh1['b'+str(k+1)].value,sh1['c'+str(k+1)].value])
frame1=pd.DataFrame(np.random.rand(i*j).reshape((j,i)))
frame0=pd.DataFrame(np.arange(j),columns=['quesnum'])+1
frame2=frame1.merge(frame0,left_index=True,right_index=True,)
for k in range(1,i+1):
    frame3=frame2.sort_values(by=k-1)
    k_file=open(r'c:\python tem\file_%s.txt'%(k),'w',encoding='utf-8')
    for l in range(1,j+1):
        ques_num=list(frame3.iloc[:,-1])[l-1]
        k_file.write('question %s:'%(l))
        k_file.write(str(question_dic[ques_num][0])+' answer:')
        k_file.write(str(question_dic[ques_num][1])+'\n')
    k_file.close()