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

def readFiles(path,num):
    rtn_list=[]
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(num)
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

def read_content(path,num):
    global active_filename
    read_list=[]
    read_list1=[]
    for filename in path:
        active_filename = os.path.split(filename)[1]
        print '==>', active_filename
        if unicode(filename, 'cp936').find(u'账号使用情况') > -1:
            print filename, u'账号使用情况'
            read_list = readFiles(filename,num)
            
            read_list.pop(0)
            
        if unicode(filename, 'cp936').find(u'直接从平台导的表格') > -1:
            print filename, u'直接从平台导出'
            read_list1 = readFiles(filename,num)
            read_list1.pop(0)
            read_list1.pop(-1)
    
    return read_list, read_list1
#--------------------------------------------------------------------------------------

def find_code_data(loginName,loginIds,list2):
    #print len(loginIds),'++++++++++++++++++++++'
    options=[]
    options1=[]
    loginId=u''
    #print loginIds
    for i in range(len(loginIds)):
        for r in range(len(list2)):
            if list2[r][0]==loginIds[i]:
                #print list2[r][0],'+',loginIds[i]
                options.append(list2[r][2:])
                break  
        if len(loginId)==0:
            loginId=loginIds[i]
        else:
            loginId+=u'/'+loginIds[i][-3:]
    rows=[]
    d=0
    e=0
    f=0
    h=0
    j=0
    l=0
    
    for r in range(len(options)):
        
       d+=float(options[r][0])
       e+=float(options[r][1])
       f+=float(options[r][2])
       h+=float(options[r][4])
       j+=float(options[r][6])
       l+=float(options[r][7])
       
    try:    
       rows.append([loginId,loginName,str('%.0f' % d),str('%.2f' % (e/len(options))),str('%.0f' % f),str('%.2f' % (f/d*100)),
                         str('%.0f' % h),str('%.2f' % (h/d*100)),str('%.0f' % j),str('%.2f' % (j/d*100)),
                         str('%.0f' % l),str('%.2f' % (l/d*100))])
    except:
       pass

       
    return rows

def data_add(log_in,list2):
    datas=[]
    #print log_in
    for key in log_in:
        items=log_in[key]
        datas+=find_code_data(key,items,list2)
   
    datas=sorted(datas, cmp=lambda y,x : cmp(float(x[2]), float(y[2]))) #排序
    
    return datas
#-------------------------------------------------------------------------------

def some_set_map(list):
    
    data_dict={}
    datas=[]
    datass=[]
    
    for i in range(len(list)):
        key=list[i][0]
        if data_dict.has_key(key):
            data_dict[key].append(list[i])
        else:
            data_dict[key]=[list[i]]

    for key in data_dict:
        item=data_dict[key]
        datas+=some_data_add(key,item)
    
    for i in range(len(datas)):
        if i%2!=0:
           datass.append(datas[i])
    datas=sorted(datas, cmp=lambda y,x : cmp(float(x[2]), float(y[2]))) #排序
    return datas
              
def some_data_add(key,item):

    
    rows=[]
    d=0#业务量
    e=0#时间
    f=0##差错数、差错率
    h=0#回收量、回收率
    j=0#修改数、修改率
    l=0#拆分数、拆分率
    for r in range(len(item)):
        #print item[r][3]
        
        d+=float(item[r][3])#业务量
        e+=float(item[r][4])#时间
        f+=float(item[r][5])#差错数、差错率
        h+=float(item[r][7])#回收量、回收率
        j+=float(item[r][9])#修改数、修改率
        l+=float(item[r][11])#拆分数、拆分率
    try:    
        rows.append([key,item[r][1],str('%.0f' % d),str('%.2f' % (e/len(item))),str('%.0f' % f),str('%.2f' % (f/d*100)),
                         str('%.0f' % h),str('%.2f' % (h/d*100)),str('%.0f' % j),str('%.2f' % (j/d*100)),
                         str('%.0f' % l),str('%.2f' % (l/d*100))])
    except:
        pass
    return rows   

 
#--------------------------------------------------------------------------------------        

        
def set_map(list,list2):#生成key=用户名,value=账号（可以有多个账号）的字典
    log_in={}
    for i in range(len(list)):
        key=list[i][1]
        if log_in.has_key(key):
           log_in[key].append(list[i][0])
        else:
           log_in[key]=[list[i][0]]


           
    #for key in log_in:
        #print key,':',log_in[key]
    return data_add(log_in,list2)

       
def fill_login(list):#Excel表格的携带 
   first_name=u''
   for i in range(len(list)):
       if list[i][1]!=u'':
           first_name= list[i][1]
       else:
            list[i][1]=first_name    
   return list
#--------------------------------------------------------------------------------------

def kill_some_account(list):
    
    list1=[]
    list2=[]
    lst1=[]
    
    for i in range(len(list)):
       
        if (list.count(list[i]))==1:
            list1.append(list[i])
        else:
            list2.append(list[i])
            
    list2=sorted(list2, key=lambda list2: list2[0][-3:])
    list2.append('wangan')
        
    for r in range(len(list2)):
        if r+1<len(list2):
            if list2[r]!=list2[r+1]:
               lst1.append(list2[r])
        
    list1.extend(lst1)
    

    return list1
    
#--------------------------------------------------------------------------------------

def to_tal(list):
    #print list[0][2]
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
        rows=[str(len(list)),str('%.0f' % d),str('%.2f' % (e/len(list))),str('%.0f' % f),str('%.2f' % (f/d*100)),
                         str('%.0f' % h),str('%.2f' % (h/d*100)),str('%.0f' % j),str('%.2f' % (j/d*100)),
                         str('%.0f' % l),str('%.2f' % (l/d*100))]
    except:
        pass
   
    return rows

#--------------------------------------------------------------------------------------

def test_data(a_list,b_list):#比较两个表中的数据是否一致
    error_list=[]
    for i in range(len(b_list)):
        isExists=False
        for j in range(len(a_list)):
            if b_list[i][0]==a_list[j][0]:
                isExists=True
                break
        if not isExists:        
            error_list.append([b_list[i][0],u'直接从平台导的表格中的操作员代码在账号使用情况表中不存在'])
            
    return error_list
            
#--------------------------------------------------------------------------------------
def doWork():
    a_data_list=[]
    b_data_list=[]

    a1_data_list=[]
    b1_data_list=[]
    
    cwd=os.getcwd()
    if os.path.isfile(cwd+'\\result.xls'):
      os.remove(cwd+'\\result.xls')
      
    path=glob.glob(cwd+'/*.xls')
    path.sort()
    
    a_data_list, b_data_list = read_content(path,0)#读取数据
    a1_data_list, b1_data_list = read_content(path,1)#读取数据
    
    #error_list=test_data(a_data_list, b_data_list)#测试
    b_data_list=some_set_map(b_data_list)
    b1_data_list=some_set_map(b1_data_list)

    a_data_list=fill_login(a_data_list)
    a1_data_list=fill_login(a1_data_list)
    
    #print len(a_data_list),len(a1_data_list),'+++++1'

    a_data_list=kill_some_account(a_data_list)
    a1_data_list=kill_some_account(a1_data_list)
    
    #print len(a_data_list),len(a1_data_list),'+++++2'
   
    totalSum=set_map(a_data_list,b_data_list);
    totalSum1=set_map(a1_data_list,b1_data_list);

    a=to_tal(totalSum)
    b=to_tal(totalSum1)
   
    write(cwd,totalSum,totalSum1,a,b,
        [u'序号',u'操作员代码',u'操作员姓名',u'处理业务量',u'平均处理\n时间（秒）',u'差错数',u'差错率',u'回收量'
         ,u'回收率',u'转录入修改数',u'转录入修改率(%)',u'转影像拆分数',u'转影像拆分率(%)',u'备注'],'result.xls')
   
    #writeReport(cwd,error_list,[u'序号',u'操作员代码',u'测试信息'])
          
                    
#--------------------------------------------------------------------------------------    

def write(cwd,totalSum,totalSum1,a,b,column_name,name):
    
    book=xlwt.Workbook(encoding='utf-8')

    font=xlwt.Font()#为样式创建字体
    fontA=xlwt.Font()#为样式创建字体
    
    alignment=xlwt.Alignment()#设置对齐方式
    alignment.horz = xlwt.Alignment.HORZ_CENTER#居中

    font.height=350#设置字体大小
    font.bold=True#设置是否黑体

    fontA.height=250#设置字体大小
    fontA.bold=True#设置是否黑体

    style=xlwt.XFStyle()#大标题样式
    center_style=xlwt.XFStyle()#居中样式
    column_style=xlwt.XFStyle()#列名样式
    
    style.font=font
    style.alignment = alignment
    
    center_style.alignment = alignment
    
    column_style.alignment = alignment
    column_style.font=fontA
    
    
    sheet=book.add_sheet('891')
    sheet.write_merge(0,0,0,13,'891账号作业明细',style)#合并单元格
    sheet.write_merge((len(totalSum)+2),(len(totalSum)+2),0,1,'合计',column_style)
    
    
    for i in range(len(column_name)):
        sheet.write(1,i,column_name[i],column_style)#写入列名
    
    for r in range(len(totalSum)):
        sheet.write(2+r,0,r+1,center_style)
        for c in range(len(totalSum[r])):
            sheet.write(2+r,c+1,totalSum[r][c],center_style)
    
    for i in range(len(a)):
        sheet.write(len(totalSum)+2,i+2,a[i],center_style)
    

    
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
        sheet1.write(len(totalSum1)+2,i+2,b[i],center_style)
    

    col_width=[2002,7212,3000,3000,4200,2800,2800,2800,2800,3730,3730,3730,3730,2800]
    for i in range(14):
        sheet.col(i).width = col_width[i]#设置列宽
        sheet1.col(i).width = col_width[i]
      
    book.save(cwd+'\\'+name)#保存
    



def writeReport(cwd,error_list,column_name):
    book=xlwt.Workbook(encoding='utf-8')
    
    sheet=book.add_sheet('sheet')
   
    for i in range(len(column_name)):
        sheet.write(1,i,column_name[i])#写入列名
    for r in range(len(error_list)):
        sheet.write(2+r,0,r+1)
        for c in range(len(error_list[r])):
            sheet.write(2+r,c+1,error_list[r][c])       
    book.save(cwd+'\\report.xls')
   
#-------------------------------------------------------------------------------------- 
              
if __name__=='__main__':
    
    doWork()
#     try:
#         doWork()
#     except:
#         print 'File Error!'
    raw_input('Please input end!!!')


