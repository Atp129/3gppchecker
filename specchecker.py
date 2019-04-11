#!/usr/bin/env python
# encoding: utf-8
# # @author: Hongkang LI

from win32com import client as wc
import re
import os

from loggingset import *
from unzip import *


def open_doc(file):
    word = wc.Dispatch('Word.Application')
    word.Visible = 0
    word.DisplayAlerts = 0
    doc = word.Documents.Open(file)
    return doc, word


def new_doc():
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Add()
    # doc.SaveAs(file, 2)
    return doc, word


class DocContent(object):
    def __init__(self, name, start, exit):
        self.name = name
        self.start_format = start
        self.exit_format = exit
        self.start = None
        self.end = None

    def is_start(self, format):
        if self.start_format in format:
            return True
        else:
            return False

    def is_end(self, next_format):
        if self.exit_format == '':
            return False
        if self.exit_format in next_format:
            return True
        else:
            return False

    def record_start(self, start):
        self.start = start

    def record_end(self, end):
        self.end = end


class Chapter(object):
    def __init__(self, title):
        self.title = title
        if 'Annex' in title:   # 文档Annex 格式
            string = title.split(':', 1)
        else:                   # 一般章节
            string = title.split('\t')
        self.id = string[0]
        self.name = string[1]
        self.start = None
        self.end = None
        info(string)

    def set_start(self, start):
        self.start = start

    def set_end(self, end):
        self.end = end


class SpecDoc(object):
    def __init__(self, file):
        self.doc, self.app = open_doc(file)
        self.id = ''
        self.chapter_list = []
        self.contents = []
        self.content_index = 0
        self.set_content()
        self.chapter_index = 0
        self.new_chapter = None
        self.EOF = False
        self.ref_list = []
        self.version = ''

    def set_content(self):
        title = DocContent('title', 'Z', 'FP')
        frontpage = DocContent('frontpage', 'FP', 'TT')
        menu = DocContent('menu', '目录', '标题')
        text = DocContent('text', '标题', 'not available')
        self.contents = [title, frontpage, menu, text]

    def scan(self):
        para_count = self.doc.Paragraphs.Count
        self.EOF = False   #  初始化文档结尾判断

        for i in range(0, para_count):
            current_para = self.doc.Paragraphs[i]  # 记录当前行
            if i < para_count - 1:  # 判断是否进行到文档尾行
                next_para = self.doc.Paragraphs[i+1]
            else:               # 如果在文档尾行，直接保存段落结尾并退出
                self.contents[self.content_index].record_end(current_para.Range.End)
                self.EOF = True

            # if current_para is None:
            #     continue
            if next_para is None:
                continue

            try:
                info(current_para.Range.Text)
                info(current_para.Range.Style.NameLocal)
                if self.contents[self.content_index].start is None:
                    if self.contents[self.content_index].is_start(current_para.Range.Style.NameLocal):
                        self.contents[self.content_index].record_start(current_para.Range.Start)
                    # 内容开始未记录时，检查是否满足开始条件，如果满足，则记录开始点
            except (Exception) as e:
                warning(e)

            if self.contents[self.content_index].name == 'title':
                self.check_name(current_para)

            if self.contents[self.content_index].name == 'reference':
                self.check_reference(current_para)

            if self.contents[self.content_index].name == 'text':
                self.check_chapter(current_para, next_para)

            try:
                if self.contents[self.content_index].is_end(next_para.Range.Style.NameLocal) or self.EOF:
                    self.contents[self.content_index].record_end(current_para.Range.End)
                    self.content_index = self.content_index + 1
                    # 检查下一行是否符合推出条件，如果满足直接退出

                debug(self.contents[self.content_index].name)

            except (Exception) as e:
                warning(e)

            # try:
            #     if self.contents[self.content_index].name == 'text':
            #         self.check_chapter(current_para, next_para)
            #     # 如果到了正文阶段，则扫描正文
            # except (Exception) as e:
            #     warning(e)
            #     if e is IndexError:
            #         warning(e)
            #         warning(self.contents)
            #         warning(self.content_index)
            #         warning(self.chapter_index)
            #     if e is NameError:
            #         warning(current_para)
            #         warning(next_para)

    def check_content(self, current, next):
        if self.contents[self.content_index].is_start(current):
            self.content_index[self.content_index].record_start

    def check_chapter(self, current, next):
        try:
            if '标题' in current.Range.Style.NameLocal:
                debug('start chapter')
                self.new_chapter = Chapter(current.Range.Text)
                self.new_chapter.set_start(current.Range.Start)
        except Exception as e:
            warning(e)

        try:
            info(self.new_chapter.name)
            if 'References' in self.new_chapter.name:
                self.check_reference(current)
        except Exception as e:
            warning(e)

        try:
            if ('标题' in next.Range.Style.NameLocal) or self.EOF:
                debug('end chapter')
                self.new_chapter.set_end(current.Range.End)
                self.chapter_list.append(self.new_chapter)
        except Exception as e:
            warning(e)

    def check_reference(self, current):
        pattern = r'\[(?P<id>\d+)\]\s+(?P<code>[^:]+):(?P<name>.+).'
        match = re.search(pattern, current.Range.Text)
        if match:
            self.ref_list.append([match.group('id'), match.group('code'), match.group('name')])
            info([match.group('id'), match.group('code'), match.group('name')])

    def check_name(self, current):
        if 'ZA' == current.Range.Style.NameLocal:
            pattern = r'3GPP TS (?P<id>[\d.]+) (?P<version>V[\d.]+)'
            match = re.search(pattern, current.Range.Text)
            if match:
                self.id = match.group('id')
                self.version = match.group('version')

    def generate(self, path):
        for chapter in self.chapter_list:
            id = re.sub(r"\x0c", '', chapter.id)
            print([chapter.id, chapter.name])
            mkdir(os.path.join(path, self.id, self.version))
            filename = id + '.html'
            info(filename)
            self.app.Selection.SetRange(chapter.start, chapter.end)
            self.app.Selection.Copy()
            doc2, app2 = new_doc()
            doc2.Content.Paste()
            doc2.SaveAs(os.path.join(path, self.id, self.version, filename), 8)
            doc2.Close()


def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        info('path exist')


def convert_path(path):
    g = os.walk(path)
    unzip_path = r'D:\unzip'

    for path, dir_list, file_list in g:
        for file_name in file_list:
            zipped_file = os.path.join(path, file_name)
            info(zipped_file)
            un_zip(zipped_file, unzip_path)


def convert_unzip(path):
    g = os.walk(path)

    for path, dir_list, file_list in g:
        for file_name in file_list:
            file = os.path.join(path, file_name)
            info(file)
            if '.doc' in file_name:
                convert_file(file)


def convert_file(file):
    path = r'C:\Users\romain.li\PycharmProjects\ghostwriter\html'
    spec_doc = SpecDoc(file)
    spec_doc.scan()
    try:
        spec_doc.generate(path)
    except Exception as e:
        warning(e)


if __name__ == '__main__':
    file = r'C:\Users\romain.li\PycharmProjects\ghostwriter\38124-f10.doc'
    doc38214 = SpecDoc(file)

    lst = []
    path = r'C:\Users\romain.li\PycharmProjects\ghostwriter\html'
    root = r''
    # for i in range(1,100):
    #     name = str(i) + ' xxxx'
    #     x = Chapter(name)
    #     print(name.split(' ',1))
    #     lst.append(x)
    #
    # for chap in lst:
    #     print(chap.id, chap.name)

    # string = '[1]	3GPP TR 21.905: "Vocabulary for 3GPP Specifications".'
    # string = '3GPP TS 38.124 V15.1.0 (2018-03)'
    # pattern = r'3GPP TS (?P<id>[\d.]+) (?P<version>V[\d.]+)'
    # # pattern = r'(?P<id>\[\d+\])\s+(?P<code>\w+):\s+(?P<name>\w+)'
    # # pattern = r'\[(?P<id>\d+)\]\s+(?P<code>.+):(?P<name>.+).'
    # match = re.search(pattern, string)
    # if match:
    #     print(match.group('id'))
    #     print(match.group('version'))
    #     # print(match.group('name'))
    # id = re.sub(r"[\W]", '_', 'Annex A (informative)')
    # print(id)

    # doc38214.scan()
    #
    # for chapter in doc38214.chapter_list:
    #     print(chapter.id, chapter.name)
    #     print(chapter.start, chapter.end)
    #
    # for ref in doc38214.ref_list:
    #     print(ref)
    #
    # try:
    #     doc38214.generate(path)
    # except Exception as e:
    #     warning(e)

    # convert_path(r'D:\spec\38series')
    unzip_path = r'D:\unzip'
    convert_unzip(unzip_path)
