#-*- coding:utf-8 -*-

import requests
from unzip import *
from specdecode import *



def filedld(url, name):
    print "downloading with requests"
    # url = 'http://ww.pythontab.com/test/demo.zip'
    r = requests.get(url)
    with open(name, "wb") as code:
        code.write(r.content)

if __name__ == "__main__":
    # url = "http://www.3gpp.org/ftp/Specs/archive/38_series/38.331/38331-f30.zip"
    # name = "d:\\38331-f30.zip"
    # filedld(url, name)
    # un_zip(name)

    saveashtml(r'D:\38331-f30.zip_files\38331-f30.doc', r'd:\38331-f30.html')