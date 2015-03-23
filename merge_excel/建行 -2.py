#-*- coding:utf8-*-
#################################################################
# Author     :An                                            
# Version    : 1.0.0.0                                         
# Date       : 2014-12-25                                      
# Description:                                                 
#################################################################
import os
import glob
from string import *
import pyExcelerator as xlwd
import re
import copy
import xlrd
import xlwt
from win32com.client import Dispatch

error_list=[]
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

def readXls(path,num):
    rtn_list=[]
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(num)
    rows = sheet.nrows
    cols = sheet.ncols
    for r in range(2,rows):
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

#-------------------------------------------------------------------------------

def some_set_map(list):

    isSome=False
    data_dict={}
    datas=[]
    
    for i in range(len(list)):
        key=list[i][2]
        for j in range(len(list)):
            if list[i][2]==list[j][2] :
                isSome=True
                break
        if isSome:
            if data_dict.has_key(key):
                data_dict[key].append(list[i])
            else:
                data_dict[key]=[list[i]]
     
    for key in data_dict:
        item=data_dict[key]
        
        if key!=u'' and not re.findall(u'[0-9]+',key):
            datas+=some_data_add(key,item)
        
    datas=sorted(datas, cmp=lambda y,x : cmp(float(x[2]), float(y[2]))) #排序
    
    return datas



def set_id(list):
    
    for i in range(len(list)):
        if len(list[i])==12:
            list.insert(0,list[i][:8])
            break
        
    for i in range(len(list)):        
        if list[i]==u'':
            list[i]=u'#'
            
        if len(list[i])==1 or len(list[i])==9:
           list[i]=u'000'+list[i]
              
        if len(list[i])==2 or len(list[i])==10: 
            list[i]=u'00'+list[i][-2:]

        if len(list[i])==3 or len(list[i])==11:
           list[i]=u'0'+list[i]
            
        if len(list[i])==12 : 
            list[i]=list[i][-4:]
            
    return list           
    
def some_data_add(key,item):
    rows=[]
    ID_dict={}
    
    d=0#业务量
    e=0#时间
    f=0##差错数、差错率
    h=0#回收量、回收率
    j=0#修改数、修改率
    l=0#拆分数、拆分率
    for r in range(len(item)):
        
        ID=item[r][1]
        
        if ID_dict.has_key(key):
                ID_dict[key].append(ID)
        else:
            ID_dict[key]=[ID]
            
        d+=float(item[r][3])#业务量
        e+=float(item[r][4])#时间
        f+=float(item[r][5])#差错数、差错率
        h+=float(item[r][7])#回收量、回收率
        j+=float(item[r][9])#修改数、修改率
        l+=float(item[r][11])#拆分数、拆分率
    
    lsts=set_id(u'/'.join(ID_dict[key]).split('/'))
    ids=lsts.pop(0)
    listId=set(lsts)
 
    try:    
        rows.append([ids+u'/'.join(listId),key,str('%.0f' % d),str('%.2f' % (e/3)),str('%.0f' % f),str('%.2f' % (f/d*100)),
                         str('%.0f' % h),str('%.2f' % (h/d*100)),str('%.0f' % j),str('%.2f' % (j/d*100)),
                         str('%.0f' % l),str('%.2f' % (l/d*100))])
    except:
        pass
    
    return rows   
    
#--------------------------------------------------------------------------------------

def to_tal(list):
    
    d=0#业务量
    e=0#时间
    f=0#差错数、差错率
    h=0#回收量、回收率
    j=0#修改数、修改率
    l=0#拆分数、拆分率
    rows=[]
    for r in range(len(list)):
        
        d+=float(list[r][2])#业务量
        e+=float(list[r][3])#时间
        f+=float(list[r][4])#差错数、差错率
        h+=float(list[r][6])#回收量、回收率
        j+=float(list[r][8])#修改数、修改率
        l+=float(list[r][10])#拆分数、拆分率

    try:    
        rows=[str(len(list)),str('%.0f' % d),str('%.2f' % (e/3)),str('%.0f' % f),str('%.2f' % (f/d*100)),
                         str('%.0f' % h),str('%.2f' % (h/d*100)),str('%.0f' % j),str('%.2f' % (j/d*100)),
                         str('%.0f' % l),str('%.2f' % (l/d*100))]
    except:
        pass
   
    return rows

#--------------------------------------------------------------------------------------
def return_data(list891,list895,cwd):
    lst1=[]
    lst2=[]

    a=[]
    b=[]
   
    lst1=some_set_map(list891)
    lst2=some_set_map(list895)

    a=to_tal(lst1)
    b=to_tal(lst2)

    write(cwd,lst1,lst2,a,b,
        [u'序号',u'操作员代码',u'操作员姓名',u'处理业务量',u'平均处理\n时间（秒）',u'差错数',u'差错率',u'回收量'
         ,u'回收率',u'转录入修改数',u'转录入修改率(%)',u'转影像拆分数',u'转影像拆分率(%)',u'备注'],'result.xls')
        
#--------------------------------------------------------------------------------------
def contents(path):

    global active_filename
    
    list891=[]
    list895=[]
    
    parents = os.listdir(path)
    
    for parent in parents:
        child = os.path.join(path,parent)
        
        if os.path.isdir(child):
           contents(child)
           
        else:
            active_filename=os.path.split(child)[1]
            print active_filename

            if re.findall(u'(.xls)$',active_filename):
                
               for i in range(len(readXls(child,0))):
                   list891.append(readXls(child,0)[i])

               for i in range(len(readXls(child,1))):
                   list895.append(readXls(child,1)[i])
             
    return_data(list891,list895,path)  
           
            
    
#--------------------------------------------------------------------------------------
def doWork():
    
    cwd=os.getcwd()
    
    if os.path.isfile(cwd+'\\result.xls'):
      os.remove(cwd+'\\result.xls')
    
    contents(cwd)
        
#--------------------------------------------------------------------------------------    

def write(cwd,totalSum,totalSum1,a,b,column_name,name):
    
    book=xlwt.Workbook(encoding='utf-8')

    font=xlwt.Font()
    font.height=350
    font.bold=True
    
    fontA=xlwt.Font()
    fontA.height=200
    fontA.bold=True

    alignment=xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER

    style=xlwt.XFStyle()
    center_style=xlwt.XFStyle()
    column_style=xlwt.XFStyle()

    borders = xlwt.Borders()
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    borders.bottom_colour=0x3A
    
    style.font=font
    style.alignment = alignment
    
    center_style.alignment = alignment
    
    column_style.alignment = alignment
    column_style.font=fontA
    column_style.borders = borders
    
    
    sheet=book.add_sheet('891')
    sheet.write_merge(0,0,0,13,'891账号作业明细',style)
    sheet.write_merge((len(totalSum)+2),(len(totalSum)+2),0,1,'合计',column_style)
    
    
    for i in range(len(column_name)):
        sheet.write(1,i,column_name[i],column_style)
    
    for r in range(len(totalSum)):
        sheet.write(2+r,0,r+1,center_style)
        for c in range(len(totalSum[r])):
            sheet.write(2+r,c+1,totalSum[r][c],center_style)
    
    for i in range(len(a)):
        sheet.write(len(totalSum)+2,i+2,a[i],column_style)
    
    
    sheet1=book.add_sheet('895')
    sheet1.write_merge(0,0,0,13,'895账号作业明细',style)#合并单元格
    sheet1.write_merge((len(totalSum1)+2),(len(totalSum1)+2),0,1,'合计',column_style)
    
    for i in range(len(column_name)):
        sheet1.write(1,i,column_name[i],column_style)#写入列名

    for r in range(len(totalSum1)):
        sheet1.write(2+r,0,r+1,center_style)
        for c in range(len(totalSum1[r])):
            sheet1.write(2+r,c+1,totalSum1[r][c],center_style)

    for i in range(len(b)):
        sheet1.write(len(totalSum1)+2,i+2,b[i],column_style)
    

    col_width=[2002,7212,3000,3000,4200,2800,2800,2800,2800,3730,3730,3730,3730,2800]
    for i in range(14):
        sheet.col(i).width = col_width[i]#设置列宽
        sheet1.col(i).width = col_width[i]
      
    book.save(cwd+'\\'+name)#保存
    

#-------------------------------------------------------------------------------------- 
              
if __name__=='__main__':
    
    doWork()

    raw_input('Please input end!!!')


