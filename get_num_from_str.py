#coding = utf8
#get num from string and write sum formula

import re

def get_formula():
    def_str = input('enter string to be cal:')
    begin_str = input('enter word before every num:')
    end_str = input('enter word after every num:')
    pat1 = '%s([\d\.]+)?%s'%(begin_str,end_str)
    num_list = re.findall(pat1,def_str)
    formula_str = ''
    sum_num = 0
    for i in num_list:
        formula_str = formula_str+i+'+'
        sum_num += float(i)
    formula_str = formula_str[:-1]+'='+str(sum_num)
    print(formula_str)


if __name__ == '__main__':
    get_formula()
