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
import xlwt

error_list=[]
except_list=[]
ts=u'image	№	日付	Q1	Q2	Q3	Q3その他	Q4	Q4その他	Q5	Q5④宿	Q5④泊	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6行き	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6山内	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6帰り	Q6そのた	Q7①	Q7②	Q7③	Q7④	Q7⑤	Q7⑥	Q7⑦	Q7その他	Q8	Q8施設名	Q8理由	Q8理由その他	Q9	Q9その他	Q10	Q10④FA	Q10⑤FA	Q10⑥FA	Q10⑦FA	Q11①	Q11②	Q11③	Q11③その他	Q12①	Q12②	Q12③	Q12④	Q12⑤	Q12⑥	Q12⑦	Q12⑧	Q12⑨	Q12⑩	Q12⑩その他	Q13	Q14①	Q14②	Q14③	Q14④	Q14⑤	Q14⑥	Q14⑦	Q14⑧.1	Q14⑧.2	Q15①	Q15②	Q15③	Q15④	Q15⑤	Q15⑥	Q15⑦	Q15⑧	Q15⑨	Q15⑨その他	Q16自由回答	性別	年齢	お住まい：滋賀	お住まい：京都市	お住まい：京都府	お住まい：それ以外'.split('\t')
ts_chu=[u'Image',u'No.',u'項目名',u'問題',u'処理方法']

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


def do_data_content(list1,maps):
    list1.pop(0)
    rtn_list=[]
    
    for i in range(len(list1)):
        
        if len(list1[i].strip())==0:
            continue
        temp_list = list1[i].split('\t')
       
        if len(temp_list)<34:
            print
            error_list.append([ '***Error*** File:'+active_filename+' Line: '+str(i+2)+' length is not 34!'])
            return None
        
        wz_list=get_value_from_map(maps,temp_list[0])
    
        if wz_list is not None:
            rtn_list.append(do_data_item(temp_list,wz_list,i+2))
                
    return rtn_list

#---------------------------------------------------------------------------------------------------------------------------------------------
def set_ma(strs,length,code,title):
    ret_list=[u'' for i in range(length)]
    if len(strs)>0:
        temp=strs.split(u'.')
        if len(temp)>length:
            error_list.append([code,title,u'选项不明',strs])
        else:       
            ret_list=initItem(temp,length)
    return ret_list

def initItem(list1,length):
    if len(list1)<length:
        list1.extend([u'' for i in range(length-len(list1))])
    return list1


def set_sa(strs,image,no,title):

    if strs==u'':
        return None
    
    if strs.find('.')==-1:
        return strs
    else:
        except_list.append([image,no,title,strs,u'多选'])
        return None
        
    
def set_fa(strs,image,no,title):
    if strs==u'':
        return None
    
    if strs!=u'+':
        return strs
    else:
        except_list.append([image,no,title,strs,u'字段为+'])
        return None


def get_effective_data(strs):
    rtn_list=[]
    while len(strs)>0:
        tt=re.findall(u'^([0-9](\.[0-9])+)',strs)
        if tt:
            rtn_list.append(tt[0][0])
            strs=strs[len(tt[0][0]):]
        else:
            rtn_list.append(strs[0])
            strs=strs[1:]
 
    return rtn_list


def split_chars(strs,length,image,no,start,flog=1):
    temp = get_effective_data(strs)

    if len(temp)>length:
        error_list.append(u'***Error*** File:'+unicode(active_filename,'cp932')+u' Image:'+image+u' Col:'+ts[start]+u' 分割长度大于'+ str(length)+u'-->'+strs)

        return [u'' for j in range(length)]

    for i in range(len(temp)):
        if temp[i]==u'0':
            temp[i]=u''
        else:
            temp[i]=get_sa_value(temp[i],image,no,ts[start+i*flog])

    return temp

def get_sa_value(strs,image,no,title):
    if strs.find(u'.')>-1:
        except_list.append([image,no,title,strs,u'多选'])
        strs=u''
        
    return strs
#--------------------------------------------------------------------------------------------


def do_data_item(szlist,wzlist,line):
    lists=[]
    #print szlist,wzlist
    lists.append(szlist[0].split('\\')[-1])
    lists.append(szlist[1].zfill(3))
    lists.append(re.findall(u'[0-9]+',wzlist[0].split('\\')[4])[0].zfill(3))
    for i in range(2,5):
        lists.append(set_sa(szlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+1]))
    lists.append(set_fa(wzlist[1],szlist[0].split('\\')[-1],szlist[1],ts[6]))
    lists.append(set_sa(szlist[5],szlist[0].split('\\')[-1],szlist[1],ts[7]))
    lists.append(set_fa(wzlist[2],szlist[0].split('\\')[-1],szlist[1],ts[8]))
    lists.append(set_sa(szlist[6],szlist[0].split('\\')[-1],szlist[1],ts[9]))
    lists.append(szlist[7].zfill(3))
    lists.append(szlist[8].zfill(3))

    lists.extend(set_ma(szlist[9],12,szlist[0].split('\\')[-1],ts[12]))
    lists.extend(set_ma(szlist[10],12,szlist[0].split('\\')[-1],ts[24]))
    lists.extend(set_ma(szlist[11],12,szlist[0].split('\\')[-1],ts[36]))

    lists.append(set_fa(wzlist[3],szlist[0].split('\\')[-1],szlist[1],ts[48]))

    for i in range(12,19):
        lists.append(set_sa(szlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+37]))

    lists.append(set_fa(wzlist[4],szlist[0].split('\\')[-1],szlist[1],ts[56]))
    lists.append(set_sa(szlist[19],szlist[0].split('\\')[-1],szlist[1],ts[57]))
    lists.append(set_fa(wzlist[5],szlist[0].split('\\')[-1],szlist[1],ts[58]))
    lists.append(set_sa(szlist[20],szlist[0].split('\\')[-1],szlist[1],ts[59]))
    lists.append(set_fa(wzlist[6],szlist[0].split('\\')[-1],szlist[1],ts[60]))
    lists.append(set_sa(szlist[21],szlist[0].split('\\')[-1],szlist[1],ts[61]))
    lists.append(set_fa(wzlist[7],szlist[0].split('\\')[-1],szlist[1],ts[62]))
    lists.append(set_sa(szlist[22],szlist[0].split('\\')[-1],szlist[1],ts[63]))

    for i in range(8,12):
        lists.append(set_fa(wzlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+56]))

    for i in range(23,26):
        lists.append(szlist[i].zfill(3))

    lists.append(set_fa(wzlist[12],szlist[0].split('\\')[-1],szlist[1],ts[71]))

    lists.extend(set_ma(szlist[26],10,szlist[0].split('\\')[-1],ts[72]))

    lists.append(set_fa(wzlist[13],szlist[0].split('\\')[-1],szlist[1],ts[82]))
    lists.append(set_sa(szlist[27],szlist[0].split('\\')[-1],szlist[1],ts[83]))

    lists.extend(split_chars(szlist[28],3,szlist[0].split('\\')[-1],szlist[1],84))
    lists.extend(split_chars(szlist[29],4,szlist[0].split('\\')[-1],szlist[1],87))
    lists.extend(split_chars(szlist[30],2,szlist[0].split('\\')[-1],szlist[1],91))

    lists.extend(set_ma(szlist[31],10,szlist[0].split('\\')[-1],ts[93]))

    for i in range(14,16):
        lists.append(set_fa(wzlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+88]))

    for i in range(32,34):
        lists.append(set_sa(szlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+72]))

    for i in range(16,20):
        lists.append(set_fa(wzlist[i],szlist[0].split('\\')[-1],szlist[1],ts[i+90]))
        
    return lists

def read_txt_wz(cwd):
    print '-------------WZ Part 1-------------'
    path=glob.glob(cwd+"/*-w/*.txt")
    path.sort()
    rtn_map={}
    for filename in path:
        Txt_filename=os.path.split(filename)[1]
        print '--->',Txt_filename
        read_txt_list =readFiles(filename)
        read_txt_list.pop(0)
        if init_txt_map(rtn_map,read_txt_list,unicode(Txt_filename,'cp932'),9) is None:
            return None
    return rtn_map

def init_txt_map(maps,content_list,filename,length):
    for item in content_list:
       if len(item.strip())==0:
           continue
       temp = item.split('\t')
       if len(temp)<length:
           print '--->',filename,temp[0],'length is not ',length
           return None
       if maps.has_key(temp[0]):
           error_list.append(u'***Error*** FileName:'+filename+u' Image:'+temp[0]+u' 在文字部分出现一条以上记录，程序默认第一次出现的记录，请确认！')
       else:
           maps[temp[0]]=temp
    return 0;

def get_value_from_map(maps,key):
   
    rtn_ss=None
    try:
        rtn_ss=maps[key]
    except:
        error_list.append(u'***Error*** Image:'+key+u' 在文字部分没有该图像所对应的记录！')
    return rtn_ss

def writeReport(cwd):
    wb = xlwd.Workbook()
    xlwd.UnicodeUtils.DEFAULT_ENCODING ='cp932'
    ws = wb.add_sheet('sheet')
    for r in range(len(error_list)):
        temp_item=error_list[r]
        ws.write(r,0,temp_item)         
    wb.save(cwd+'\\report.xls')


def writeFormat(cwd,list):
    try:
        book = xlrd.open_workbook(cwd+'/Model.xls',formatting_info=True,on_demand=True)
    except:
        print '***Error*** Can\'t find Model.xls-----------'
        return

    book.sheet_by_index(0)
    wb = copy(book)
    for i in range(len(list)):
        temp=list[i]
        for j in range(len(temp)):
            wb.get_sheet(0).write(i+2,j,temp[j])
    wb.save(cwd+'/Result.xls')

def doWork():
    global active_filename
    cwd=os.getcwd()
    content_map = read_txt_wz(cwd)
    if content_map is None:
        return;
    path=glob.glob(cwd+"/*-s/*.txt")
    path.sort()
    isWrite=True
    writlist=[]
    print '-------------SZ Part 2-------------'
    for filename in path:
        active_filename=os.path.split(filename)[1]
        print '--->',active_filename
        readList = readFiles(filename)
        
        tt = do_data_content(readList,content_map)
        
        if tt:
            writlist.extend(tt)
        else:
            isWrite=False
            break
     
    if isWrite and writlist:
        write_data(cwd,writlist,ts)
    if except_list:
        write(cwd,except_list,ts_chu,'list.xls')
    if error_list:
        writeReport(cwd)
    
def write_data(cwd,list1,title):
    
    book=xlwt.Workbook(encoding='utf-8')
    
    sheet=book.add_sheet('sheet')
    
    for i in range(len(title)):
        sheet.write(0,i,title[i])
    for r in range(len(list1)):
        for c in range(len(list1[r])):
            sheet.write(1+r,c,list1[r][c])          
    																																																																																																					         
    book.save(cwd+'/Result.xls')
    
if __name__=='__main__':

    doWork()

    '''
    try:
        doWork()
    except:
        print 'File Error!'
    '''
    
    raw_input('Please input end!!!')
