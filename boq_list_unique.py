import openpyxl
import numpy as np
import pandas as pd
import re
import pprint
file_route = input('please enter file route:')
wb = openpyxl.load_workbook(file_route)
sh = wb.get_sheet_by_name('01')
list_source = []
list_dir = []
sub_word = ''
f = open(r'c:\python tem\boq_list.txt','w',encoding='utf-8')
l = open(r'c:\python tem\boq_list_noindex.txt','w',encoding='utf-8')
for i in range(1,1741):
    list_source.append(str(sh['a'+str(i)].value))
for i in range(len(list_source)):
    pat = r'\d{3}-.+'
    if re.match(pat,list_source[i]):
        sub_word=list_source[i]
        list_dir.append(sub_word)
    else:
        list_dir.append(sub_word+list_source[i])
    f.write(str(i+1)+':'+list_dir[i]+'\n')
    l.write(list_dir[i]+'\n')
