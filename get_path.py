#get path
import os

def getpath(filename):
    path = input('please enter %s path:'%(filename))
    if os.path.isfile(path):
        return path
    else:
        getpath(filename)

def getdir(dirname):
    dir = input('please enter %s dir:'%(dirname))
    if os.path.isdir(dir):
        return dir
    else:
        getdir(dirname)

