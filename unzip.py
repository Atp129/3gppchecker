# -*- coding: utf-8 -*-


import zipfile
import os


def un_zip(file_name, path):
    """unzip zip file"""
    zip_file = zipfile.ZipFile(file_name)
    if os.path.isdir(path):
        pass
    else:
        os.mkdir(path)
    for names in zip_file.namelist():
        zip_file.extract(names, path)
    zip_file.close()

