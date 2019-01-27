# -*- coding: utf-8 -*-

from win32com import client as wc
import re
import chardet

def saveashtml(file, targetfile):

    # wdFormatDocument = 0
    # wdFormatDocument97 = 0
    # wdFormatDocumentDefault = 16
    # wdFormatDOSText = 4
    # wdFormatDOSTextLineBreaks = 5
    # wdFormatEncodedText = 7
    # wdFormatFilteredHTML = 10
    # wdFormatFlatXML = 19
    # wdFormatFlatXMLMacroEnabled = 20
    # wdFormatFlatXMLTemplate = 21
    # wdFormatFlatXMLTemplateMacroEnabled = 22
    # wdFormatHTML = 8
    # wdFormatPDF = 17
    # wdFormatRTF = 6
    # wdFormatTemplate = 1
    # wdFormatTemplate97 = 1
    # wdFormatText = 2
    # wdFormatTextLineBreaks = 3
    # wdFormatUnicodeText = 7
    # wdFormatWebArchive = 9
    # wdFormatXML = 11
    # wdFormatXMLDocument = 12
    # wdFormatXMLDocumentMacroEnabled = 13
    # wdFormatXMLTemplate = 14
    # wdFormatXMLTemplateMacroEnabled = 15
    # wdFormatXPS = 18

    # file = '/FilePath/test.docx'
    # targetfile = '/DestPath/test.pdf'
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(file)
    doc.SaveAs(targetfile, 8) #17对应于下表中的pdf文件
    doc.Close()
    word.Quit()


def opendoc(file):
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(file)
    return doc, word


def newdoc(file):
    # 检查是否存在该文件
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open()
    doc.SaveAs(file)
    return doc, word


def lstSection(doc):
    snum = doc.Sections.Count

    for i in range(0, snum):
        print doc.Sections[i].Index
        para = doc.Sections[i].Range.Paragraphs
        for j in range(0, para.Count):
            # print para[j].Style.Description
            print para[j].ID


def lstBookmarks(doc):
    print doc.Bookmarks.Count
    for i in range(0, doc.Bookmarks.Count):
        print doc.Bookmarks[i].Range.Text


def lstContent(doc):
    print doc.Content.Count
    # for i in range(0, doc.Content.Count):
    #     print doc.Content[i].Range.Text


def lstCaptionLabels(doc):
    print doc.HyperLinks.Count
    # for i in range(0, doc.Content.Count):
    #     print doc.Content[i].Range.Text


def findword(App, word):
    find = App.Selection.Find
    find.Text = word
    find.Forward = True
    find.Execute()
    pos = App.Selection.Start
    App.Selection.Start = 0
    App.Selection.End = 0
    return pos


class specDoc(object):
    def __init__(self, doc, app):
        self.doc = doc
        self.app = app
        self.paralist = []
        self.scanstat = 'menu'
        self.menuStart = -1
        self.menuEnd = -1
        self.refStart = -1
        self.refEnd = -1
        # self.name = self.doc.split('\\')
        self.name = self.doc.Name
        self.chapterlist = []
        self.currentChapterTitle = []
        self.contentStart = -1
        self.contentEnd = -1
        self.titleFinish = True
        self.lastPara = 'text'

    def scan(self):
        ParaCount = self.doc.Paragraphs.Count

        for i in range(0, ParaCount):
            try:
                print self.doc.Paragraphs[i].Range.Text
                # print self.doc.Paragraphs[i].Range.Start
                # print self.doc.Paragraphs[i].Range.End
                print self.doc.Paragraphs[i].Range.Style.NameLocal

                # print self.scanstat

                docRange = self.doc.Paragraphs[i].Range

                if self.scanstat == 'menu':
                    self.getMenu(docRange)
                elif self.scanstat == 'reference':
                    self.getRef(docRange)
                elif self.scanstat == 'content':
                    self.getContent(docRange)
            except:
                pass

    def createMenu(self):
        file = self.name[0:-4] + 'Menu'
        print file
        self.createdoc(file, self.menuStart, self.menuEnd)

    def splitdoc(self):
        self.createdoc()

    def getMenu(self, range):
        if 'TOC' in range.Style.NameLocal:
            if self.menuStart == -1:
                self.menuStart = range.Start
            self.menuEnd = range.End
        elif self.menuStart != -1:
            self.scanstat = 'reference'
        return self.menuStart, self.menuEnd

    def getRef(self, range):
        if 'Guidance' in range.Style.NameLocal:
            if self.refStart == -1:
                self.refStart =range.Start
            self.refEnd = self.refEnd
        elif self.refStart != -1:
            self.scanstat = 'content'
        return self.refStart, self.refEnd

    def getContent(self, range):
        a = range.Style.NameLocal

        if u'标题' in a:
            if self.lastPara == 'text':
                self.recordChapter()
                self.contentStart = range.Start
            self.lastPara = 'title'
            self.currentChapterTitle.append(range.Text)
        else:
            self.lastPara = 'text'
            self.contentEnd = range.End

    def recordChapter(self):
        newChapter = chapter()
        newChapter.Start = self.contentStart
        newChapter.End = self.contentEnd
        newChapter.titlelist = self.currentChapterTitle
        self.chapterlist.append(newChapter)
        self.contentStart = -1
        self.contentEnd = -1
        self.currentChapterTitle = []

    def createContent(self):
        for item in self.chapterlist:
            for title in item.titlelist:
                filename = self.name[0:-4] + get_chapter_id(title) + '.html'
                self.createdoc(filename, item.Start, item.End)

    def getHC(self):
        pass

    def createdoc(self, file, start, end):
        self.app.Selection.SetRange(start, end)
        self.app.Selection.Copy()
        doc2 = self.app.Documents.Add()
        doc2.Content.Paste()
        doc2.SaveAs(file, 8)
        doc.Close()


class chapter(object):
    pass


def check_contain_chinese(check_str):
    for ch in check_str.decode('utf-8'):
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


def get_chapter_id(title):
    result = re.search(r'(?P<id>[\.\d]+)\s', title)
    if result:
        return result.group('id')
    return None

if __name__ == "__main__":
    file = r'D:\38124-f10\38124-f10.doc'
    # file = r'D:\38124-f10\38124-f10-content.doc'
    doc, App = opendoc(file)
    # doc2, App2 = newdoc(file2)
    # lstSection(doc)
    # print App.Selection.Start
    # App.Selection.Start = 0
    # App.Selection.End = 0
    # print App.Selection.End
    # print "find result"
    # print findword(App, "Contents")
    # print findword(App, "Forward")
    # start = findword(App, "Content")
    # end = findword(App, "Forward")
    # print doc.Range(start, end).Select()
    # print doc.Range(start, end).Text
    # print doc.Range(start, end).Text

    doc38124 = specDoc(doc, App)
    doc38124.scan()
    print doc38124.refEnd
    print doc38124.menuEnd
    print doc38124.refStart
    print doc38124.chapterlist
    # doc38124.createMenu()
    doc38124.createContent()




    # print doc38331.menuStart
    # print doc38331.menuEnd


    # App.Selection.SetRange(1649, 73584)
    # App.Selection.Copy()






    # lstCaptionLabels(doc)
    # lstContent(doc)

