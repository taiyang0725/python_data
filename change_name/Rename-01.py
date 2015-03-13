#-*- coding:utf8-*-
import os,shutil
import glob
from string import *
import xlrd
import pyExcelerator as xlwd

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


def get_name(cwd):
    global active_filename
    lsts=[]
    
    for folder in os.listdir(cwd):
        if os.path.isdir(folder):
            path=glob.glob(cwd+'/'+folder+'/*.*')
            path.sort()
            for filename in path:
                active_filename=filename.split('\\')[-1]
                lsts.append(filename)
                                
    return lsts
def doWork():
    name_list=[]
    cwd=os.getcwd()
   
    try:
       os.makedirs(cwd+os.sep+'new')
    except:
        pass
    
    lists=get_name(cwd)
    l=len(lists)/4
    
    for i in range(l):
        name_list.append(lists[(len(lists)*i)/l:(len(lists)*(i+1))/l])

    #print os.sep.join(name_list[0][0].split('\\')[:-1])
    #print cwd+os.sep+'new'
    '''
    
    




    '''
    for i in range(len(name_list)):     
        os.rename(name_list[i][0], cwd+os.sep+'new'+os.sep+name_list[i][3].split('\\')[-1])#修改文件或文件夹名
        os.rename(name_list[i][1], cwd+os.sep+'new'+os.sep+name_list[i][2].split('\\')[-1])
        os.rename(name_list[i][2], cwd+os.sep+'new'+os.sep+name_list[i][1].split('\\')[-1])
        os.rename(name_list[i][3], cwd+os.sep+'new'+os.sep+name_list[i][0].split('\\')[-1])
        
    shutil.rmtree(os.sep.join(name_list[0][0].split('\\')[:-1]))#删除文件夹以及文件夹的所有内容
    os.rename(cwd+os.sep+'new',os.sep.join(name_list[0][0].split('\\')[:-1]))
    
          
if __name__=='__main__':
   
    try:
        doWork()
    except:
        print 'File Error!'
    
    raw_input('Please input end!!!')
    
