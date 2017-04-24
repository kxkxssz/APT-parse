# -*- coding: utf-8 -*-：
import zipfile
import os
import time
import datetime
from xml.etree import ElementTree

#说明：对office2007及以后的版本有效，可以解析一般的xlsx,docx,pptx文件，对打开权限加密的officce文档没有解析能力。

#获取创建时间和最后修改时间
def gettime(file):
    mtime = time.ctime(os.path.getmtime(file))
    ctime = time.ctime(os.path.getctime(file))
    print(mtime)#文件的修改时间
    print(ctime)#文件的创建时间
    return mtime,ctime

def docuParse(file):#解析office文档
    a = "null"#初始化参数，避免空值出现报错
    b = "null"
    c = "null"
    d = "null"
    e = "null"
    f = "null"
    z = zipfile.ZipFile(file,mode='r')
    nl = z.namelist()
    if("ppt/" in nl):
        f = "pptx"
    elif("word/" in nl):
        f = "docx"
    elif("xl/" in nl):
        f = "xlsx"
    else: 
        print("解析失败，非符合要求的office文档")#不从文件名判断文档类型
    zcore = z.open('docProps/core.xml','r')
    zapp = z.open('docProps/app.xml','r')
    core = ElementTree.parse(zcore)
    ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
      'dc': 'http://purl.org/dc/elements/1.1/'}
    node_creator = core.findall('./dc:creator',ns)
    node_lastModifiedBy = core.findall('./cp:lastModifiedBy',ns)
    node_revision = core.findall('./cp:revision',ns)

    for creator in node_creator:
        a = creator.text#创建者
    for lastModifiedBy in node_lastModifiedBy:
        b = lastModifiedBy.text#修改者
    for revision in node_revision:
        c = revision.text#修订版本

    app = ElementTree.parse(zapp)
    ns2 = {'vt':'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
      'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
    node_application = app.findall("xmlns:Application",ns2)
    node_company = app.findall('xmlns:Company',ns2)
    for application in node_application:
        d = application.text#编辑应用
    for company in node_company:
        e = company.text#单位
    return f,a,b,c,d,e



#t = input("please input filename:")
t = "test.xlsx"
result = docuParse(t)+gettime(t)
print(result)
#输出列表格式：
#文档类型，创建者，最后修改者，修订版本，编辑应用，单位组织，文档创建时间，文档最后修改时间