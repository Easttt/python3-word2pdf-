#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
V2.0 增加GUI
'''
from comtypes.client import CreateObject
import os

fPath = input('请输入文件夹路径：')
'''
def list_nohidden(path):
    filesLieBiao = os.listdir(path)
    fileSingleList = [f for f in filesLieBiao]
    fileSingle = str(fileSingleList)
    print(fileSingle)
    judge = str.startswith(fileSingle,0,1)
    if judge == '~':
        pass
    else :
        yield fileSingle
'''
try:
    class pdfConverter:
        def __init__(self):
            #word文档转化为pdf文档时使用的格式为17
            self.wdFormatPDF = 17
            self.wdToPDF = CreateObject("Word.Application")


        def wd_to_pdf(self, folder):
            #获取指定目录下面的所有文件
            files = os.listdir(folder)
            # files = list_nohidden(folder)
            print(files)
            #获取word类型的文件放到一个列表里面
            wdfiles = [f for f in files if f.endswith((".doc", ".docx"))]
            #去除word生成的隐藏文件
            wdfiles2 = [f for f in wdfiles if not f.startswith('~') ]
            for wdfile in wdfiles2:
                #将word文件放到指定的路径下面
                wdPath = os.path.join(folder, wdfile)
                #设置将要存放pdf文件的路径
                pdfPath = wdPath.split(".")[0] + '.pdf'
                #判断是否已经存在对应的pdf文件，如果不存在就加入到存放pdf的路径内
                if pdfPath[-3:] != 'pdf':
                    pdfPath = pdfPath + ".pdf"
                #将word文档转化为pdf文件，先打开word所在路径文件，然后在处理后保存pdf文件，最后关闭
                pdfCreate = self.wdToPDF.Documents.Open(wdPath)
                pdfCreate.SaveAs(pdfPath, self.wdFormatPDF)
                pdfCreate.Close()

    if __name__ == "__main__":
        converter = pdfConverter()
        # converter.wd_to_pdf(r'C:\Users\lenovo\Desktop\yourFolder')
        converter.wd_to_pdf(fPath)
except Exception as e:
    print(e)
input("Press enter to end!") 
