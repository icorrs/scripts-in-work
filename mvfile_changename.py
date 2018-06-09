#the mv file downloaded from web always has name begin with 'MV',
#this script find the files begin with 'MV' and rename it

import os
import re

def get_path():
    '''get path the files in'''
    path1 = input('please enter path the mv files in:')
    if os.path.isdir(path1):
        return path1
    else:
        get_path()


def mvfile_rename():
    '''main attribute'''
    path1 = get_path()
    os.chdir(path1)
    list1 = os.listdir(path1)
    pat1 = r'(MV)[A-Z\s](.*)\.(.+)'
    for f in list1:
        if re.match(pat1,f):
            old_name = os.path.join(path1,f)
            new_name = os.path.join(path1,f[2:])
            os.rename(old_name,new_name)
            print('old name:%s; rename as:%s'%(f,f[2:]))
        else:
            pass


if __name__ == '__main__':
    mvfile_rename()

