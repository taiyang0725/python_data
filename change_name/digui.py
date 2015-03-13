#-*- coding:utf8-*-
import os
import glob
from string import *
import xlrd
import pyExcelerator as xlwd
import re

report_list=[]

def log(s): 
    f = open( "./report_log.rpt","a+" ,os.O_CREAT|os.O_TRUNC)
    f.write( s + '\n' )
    f.close()

def to_unicode( s ):
    rtn = u''
    try:
        rtn = unicode(s,'cp932')
    except:
        try:
            rtn = unicode(s,'utf8')
        except:
            try:
                rtn = unicode(s,'utf16')
            except:
                return unicode(s,'cp936')
    return rtn

def readXls(path):
    #print 'read..............'
    book = xlrd.open_workbook(path)
    #print 'readXls........'
    #sheet_index = book.nsheets
    #for i in range(sheet_index):
    sheet = book.sheet_by_index(0)
    rows = sheet.nrows
    cols = sheet.ncols
    ret_list=[]
    for r in range(rows):
        row = []
        for c in range(cols):
            val = sheet.cell_value(r,c)
            if not isinstance(val,unicode):
                val = str(sheet.cell_value(r,c)).strip(' ')
                val = to_unicode(val)
            val = getstr(val)
            row.append(val)
        ret_list.append(row)
    return ret_list

def write(cwd,list1,table_title,name):
    wb = xlwd.Workbook()
    xlwd.UnicodeUtils.DEFAULT_ENCODING ='cp932'
    ws = wb.add_sheet('sheet')
    for i in range(len(table_title)):
            ws.write(0,i,table_title[i])
        
    for r in range(len(list1)):
        temp_item = list1[r]
        for c in range(len(temp_item)):
                temp=temp_item[c]
                ws.write(r+1,c,temp)
    wb.save(cwd+'\\'+name)

def getstr(strs):
    tt = strs.split(u'.')
    if len(tt)==2:
       if tt[1]==u'0':
           return tt[0]
       else:
           return strs
    else:
        return strs

def writeCSV(list1,filename):
    f=open(filename,'w+',os.O_CREAT|os.O_TRUNC)
    #list1.pop(0)
    count=0
    line=0
    for i in range(len(list1)):
        item = list1[i]
        item=item[1:]
        str1=u','.join(item)
        f.write(str1.encode('cp932'))
        f.write('\n')
        line+=1
        count+=len(item)
    if count>0:
        count-=len(list1)
    report_list.append([unicode(filename.split('\\')[-1],'cp932'),line,count])
    f.close()
def contents(path):
    readList=[]
   
    parents = os.listdir(path)
    for parent in parents:
        child = os.path.join(path,parent)
        #print(child)
        if os.path.isdir(child):
           contents(child)
    
        else:
            active_filename=os.path.split(child)[1]
            print active_filename
            if child.find('.xls')>0:
               readList = readXls(child)
               
               if not re.findall(u'-list',active_filename):
                   #print child.replace('.xls','.csv') 
                   writeCSV(readList,child.replace('.xls','.csv'))
                   os.remove(child)
     
    
def doWork():
    global active_filename
    cwd=os.getcwd()

    contents(cwd)
    
    write(cwd,report_list,[u'Filename',u'Line',u'Field'],'ReportList.xls')              
                
               
if __name__=='__main__':

    doWork()

    '''
    try:
        doWork()
    except:
        print 'File Error!'
    '''    
    raw_input('Please input end!!!')
