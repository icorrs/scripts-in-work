#some file in work is printed by double-sided,in this case,the scaned pdf can't be used directly;
#the scaned image files are puted in two filedir:01 save the front side,02 sasve the back side.
#this script add dirname to the image saved in the dir,so pdf can join images as the filename.

import os
import re

from PIL import Image


def get_path(fileneeded):
    '''get image dirpath and if image need to be rotate'''
    path = input('please enter route where %s saved :'%(fileneeded))
    rota = input('please enter if image need to be rotate(0 for not, 1 for need):')
    if os.path.isdir(path) and rota in ('0','1'):
        return (path,rota)
    else:
        get_path(fileneeded)


def rotateimg(imgdir,angle):
    '''rotate the image in imgdir angle degree'''
    list1 = os.listdir(imgdir)
    for f in list1:
        path_f = os.path.join(imgdir,f)
        im = Image.open(path_f)
        im.rotate(angle).save(path_f)


def change_name(path0,reverse):
    '''change image file name'''
    list1 = os.listdir(path0)
    os.chdir(path0)
    pat = r'(\d+)_页面_(\d+)'
    if reverse == '0':
        for f in list1:
            base_name = os.path.splitext(f)[0]
            papernum = re.match(pat,base_name).group(2)
            os.rename(f,papernum+f)
    else:
        for f in list1:
            base_name = os.path.splitext(f)[0]
            papernum = len(list1)+1-int(re.match(pat,base_name).group(2))
            if papernum>9:
                os.rename(f,str(papernum)+f)
            else:
                os.rename(f,'0'+str(papernum)+f)
        

def scanjoin():
    '''main attribute'''
    path1,rota1 = get_path('front page')
    if rota1 == '0':
        change_name(path1,'0')
    else:
        rotateimg(path1,180)
        change_name(path1,'0')
    path2,rota2 = get_path('back page')
    if rota2 == '0':
        change_name(path2,'1')
    else:
        rotateimg(path2,180)
        change_name(path2,'1')


if __name__ == '__main__':
    scanjoin()
