#image compress
import os

from PIL import Image

import get_path

def get_rate():
    rate = float(input('please enter compress rate:'))
    if rate>0 and rate<1:
        return rate
    else:
        get_rate()


def imagecompress():
    path = get_path.getdir('图片所在路径')
    rate = get_rate()
    list1 = os.listdir(path)
    os.chdir(path)
    for file in list1:
        if os.path.splitext(file)[1] in ('.jpg','.jpeg','.bmp'):
            image_file = Image.open(file)
            (x,y) = image_file.size
            image_file_out = image_file.resize((int(x*rate),int(y*rate)))
            image_file_out.save(fp=('com_%s.jpg'%(file)))
        else:
            pass
    
if __name__ == '__main__':
    imagecompress()
