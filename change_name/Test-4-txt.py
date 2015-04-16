#-*- coding:utf8-*-
#################################################################
# Author     :An
# Version    : 1.0.0.1
# Date       : 2015-3-18
# Description: 
#################################################################
import os
import glob
from string import *
import pyExcelerator as xlwd
import re
import xlrd


error_list=[]
except_list=[]

def to_unicode( s ):
    rtn = u''
    try:
        rtn = unicode(s,'utf8')
    except:
        try:
            rtn = unicode(s,'cp932')
        except:
            try:
                rtn = unicode(s,'utf16')
            except:
                return unicode(s,'cp936')
    return rtn

def readFiles(filepath,code=1):
    sz = os.stat(filepath).st_size
    ifn = open(filepath,'rb')
    count = 0
    buf = ifn.read( sz )
    ifn.close
    if code==1:
        buf=to_unicode(buf)
    else:
        buf=unicode(buf,'utf16')
    ret_list=buf.split(u'\r\n')
    while len(ret_list[-1])==0:
        ret_list.pop(-1)
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

#---------------------------------------------------------------------------------------------------------------------------------------------

def getColumn(index):
    if (index/26)==0:
        return unichr(65+index)
    mod = index%26
    return unichr(65+index/26-1)+unichr(65+mod)

def get_col_value(value):
    if value is None:
        return u''
    if isinstance(value,float):
        value = getstr(str(value))
    if not isinstance(value,unicode):
        value = str(value).strip(' ')
        value = to_unicode(value)
    return value

def readExcel(path):
    rtn_list=[]
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = sheet.nrows
    cols = sheet.ncols
    for r in range(rows):
        row = []
        for c in range(cols):
            val = sheet.cell_value(r,c)
            val =get_col_value(val)
            row.append(val)
        rtn_list.append(row)
    return rtn_list

def getstr(strs):
    tt = strs.split(u'.')
    if len(tt)==2:
       if tt[1]==u'0':
           return tt[0]
       else:
           return strs
    else:
        return strs
    
#---------------读txt---------------------

def get_flog_value(flogMaps,strs):
    keys=flogMaps.keys();
    keys.sort()
    for key in keys:
        item = key.split(u':')
        if strs>=item[0] and strs<=item[1]:
            return flogMaps[key]
    error_list.append(u'***Error*** Filename:'+unicode(active_filename,'cp932')+u' Key:'+strs+u' 在base.xls里找不到相应的值,程序默认导出空！')
    return u''
    

#---------------------------------------------------------------------------------------------------------------------------------------------

def read_txt_content(path,title,length):
    print u'-------------'+title+u'-------------'
    path=glob.glob(path)
    path.sort()
    rtn_map={}
    for filename in path:
        Txt_filename=os.path.split(filename)[1]
        print '--->',Txt_filename
        read_txt_list =readFiles(filename)
        read_txt_list.pop(0)
        if init_txt_map(rtn_map,read_txt_list,unicode(Txt_filename,'cp932'),length,title) is None:
            return None
    return rtn_map

def init_txt_map(maps,content_list,filename,length,title):
    for item in content_list:
       if len(item.strip())==0:
           continue
       temp = item.split('\t')
       if len(temp)<length:
           print '--->',filename,temp[0],'length is not ',length
           return None
       if maps.has_key(temp[0]):
           error_list.append(u'***Error***<'+title+u'> FileName:'+filename+u' Image:'+temp[0]+u' 出现一条以上记录，程序默认第一次出现的记录，请确认！')
       else:
           maps[temp[0]]=temp
    return 0;

def get_value_from_map(maps,key,title):
    rtn_ss=None
    try:
        rtn_ss=maps[key]
    except:
        error_list.append(u'***Error*** <'+title+u'>Image:'+key+u' 没有该图像所对应的记录！')
    return rtn_ss


#---------------------------------------------------------------------------------------------------------------------------------------------


def create_content_data(data1,data2,data3,data4,list):
    rtn_list=[]
    rkey=[]
    rkey.extend(set(data1.keys()+data2.keys()+data3.keys()+data4.keys()))
    rkey.sort()
    for key in rkey:
        t1=get_value_from_map(data1,key,u'Part-s')
        t2=get_value_from_map(data2,key,u'Part-w')
        t3=get_value_from_map(data3,key,u'Part-zs')
        t4=get_value_from_map(data3,key,u'Part-zw')
        if t1 and t2 and t3 and t4:
            rtn_list.append(do_data_item(t1,t2,t3,t4,list))
    return rtn_list

def do_data_item(list1,list2,list3,list4,list):
    rtn_list=[]
    print list1,list2,list3,list4
    
    return rtn_list

def doWork():
    
    global active_filename
    cwd=os.getcwd()
    
    path=glob.glob(cwd+os.sep+'base'+os.sep+'*.xls')
    path.sort()
    
    for filename in path:
        active_filename=os.path.split(filename)[1]
        
        print '----->',active_filename
        readList=readExcel(filename) 
    
    content_map1=read_txt_content(cwd+'/-s/*.txt',u'part-s',52)
    if content_map1 is None:
        return
                
    content_map2=read_txt_content(cwd+'/-w/*.txt',u'Part-w',3)
    if content_map2 is None:
        return
                
    content_map3=read_txt_content(cwd+'/-zs/*.txt',u'Part-m',7)
    if content_map3 is None:
        return
                
    content_map4=read_txt_content(cwd+'/-zw/*.txt',u'Part-m',2)
    if content_map3 is None:
        return
       
    create_content_data(content_map1,content_map2,content_map3,content_map4,readList)

    '''            
    writlist=[]
    if writlist:
        write(cwd,writlist,globa_title_list,'result.xls')
    if except_list:
        write(cwd,except_list,[u'image',u'No.',u'項目名',u'問題',u'処理方法'],'list.xls')
    if error_list:
        writeReport(cwd)
    '''
        
if __name__=='__main__':

    doWork()

    '''
    try:
        doWork()
    except:
        print 'File Error!'
    raw_input('Please input end!!!')
    '''

