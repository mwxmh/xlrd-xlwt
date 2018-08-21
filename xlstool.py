#coding:utf-8
import xlrd,xlwt,sys,os,operator
from xlutils.copy import copy 
reload(sys)
sys.setdefaultencoding('utf-8')
def groupbycode():  
   bk=xlrd.open_workbook('D:\\deviceout\\pre_deviceout.xls')
   sh=bk.sheet_by_index(1)
   nrows=sh.nrows
   ncols=sh.ncols
   list=[]
   app={}
   ##从excel中读取相关资料
   for rownum in range(4,nrows):    ##从第5行开始读取
      rowvalue=sh.row_values(rownum)      
      app['id']=rowvalue[4]       ##总序號
      app['grp']=rowvalue[6].strip()       ##出货部门
      app['outintype']=rowvalue[8].strip() ##区外进口方式
      app['clearance']=rowvalue[7].strip() ##清关类型
      app['injz']=rowvalue[9].strip()      ##进口监管证件
      app['isold']=rowvalue[10].strip()    ##是否涉及旧机电
      app['isrejg']=rowvalue[11].strip()   ##是否解除监管
      app['inzcsx']=rowvalue[12].strip()   ##进口方资产属性
      app['iszj']=rowvalue[13].strip()     ##是否中检
      app['enrolno']=rowvalue[14].strip()  ##账册
      app['origsign']=rowvalue[69].strip() ##原进口流转标志
      app['isdepreciate']=rowvalue[31].strip()##是否折旧
      app['partno']=rowvalue[21]   ##料号
      app['ordergrp']=rowvalue[38].strip()   ##异动单代码
      app['paytype']=rowvalue[29].strip()   ##付费方式
      app['iscontainc']=rowvalue[47].strip() ##是否含磁
      app['groupid']=0             ##分单份数
      list.append(app.copy())            

   num=1
   for i in range(len(list)):
      count=1                      ##计数器
      partlist=[]                     
      compvalue=list[i]
      if compvalue['groupid']!=0: continue
      for j in range(1,len(list)):
         compvalue['groupid']=num
         if list[j]['groupid']!=0:continue        
         if compvalue['grp']==list[j]['grp'] and \
            compvalue['outintype']==list[j]['outintype'] and \
            compvalue['injz']==list[j]['injz'] and \
            compvalue['isold']==list[j]['isold'] and \
            compvalue['isrejg']==list[j]['isrejg'] and \
            compvalue['inzcsx']==list[j]['inzcsx'] and \
            compvalue['clearance']==list[j]['clearance'] and \
            compvalue['ordergrp']==list[j]['ordergrp'] and \
            compvalue['enrolno']==list[j]['enrolno'] and \
            compvalue['iscontainc']==list[j]['iscontainc']:      ##是否含磁是一线出口的分单匹配条件
            if (compvalue['enrolno']=='D01')or(compvalue['enrolno']=='D18'):
               if compvalue['clearance']=='繞港': 
                  if compvalue['iszj']==list[j]['iszj'] and\
                     (compvalue['origsign']==list[j]['origsign'] or ('S\U'.find(compvalue['origsign'])>=0 and 'S\U'.find(list[j]['origsign'])>=0))and\
                     (compvalue['paytype']==list[j]['paytype'] or ('Common2'.find(compvalue['paytype'])==0 and 'Common2'.find(list[j]['paytype'])==0)): ##原进口流转标志R单独分,S/U单独分；是否中检      
                     list[j]['groupid']=num
               elif compvalue['clearance']=='兩單一審':
                  if compvalue['isdepreciate']==list[j]['isdepreciate']: #两单一审时需判断是否折旧  
                     list[j]['groupid']=num    
               else:
                  list[j]['groupid']=num
            else:
               if compvalue['enrolno']=='D02': 
                  ##D02账册下料号不得超过15个              
                  if compvalue['partno']!=list[j]['partno']:
                     if partlist==[]:
                        partlist.append(compvalue['partno'])
                     if list[j]['partno']in partlist:
                        pass
                     else:
                        partlist.append(list[j]['partno'])      
                        count=count+1
                        if count==16:break
    
                  if compvalue['clearance']=='繞港':
                     if compvalue['iszj']==list[j]['iszj'] and\
                        (compvalue['origsign']==list[j]['origsign']or ('S\U'.find(compvalue['origsign'])>=0 and 'S\U'.find(list[j]['origsign'])>=0)) and\
                        (compvalue['paytype']==list[j]['paytype'] or ('Common2'.find(compvalue['paytype'])==0 and 'Common2'.find(list[j]['paytype'])==0)):                                                
                        list[j]['groupid']=num
                     else:
                        count=count-1   
                  elif compvalue['clearance']=='兩單一審': 
                     if compvalue['isdepreciate']==list[j]['isdepreciate']: #两单一审时需判断是否折旧  
                        list[j]['groupid']=num 
                     else:
                        count=count-1                                                                
                  else:
                     list[j]['groupid']=num         
                     
      num=num+1
   '''
   for ibk in list:
      if ibk['groupid'] in[11,12]:
         print ibk['id'],ibk['groupid'],ibk['partno'],ibk['origsign'],ibk['enrolno'],ibk['outintype'],ibk['clearance'],ibk['isdepreciate']
   ''' 
   writexl=xlwt.Workbook()
   sheet1=writexl.add_sheet(u'sheet1')
    
   for i in range(len(list)):
      sheet1.write(i,0,list[i]['id'])
      sheet1.write(i,1,list[i]['groupid'])
      sheet1.write(i,2,list[i]['enrolno'])
      sheet1.write(i,3,list[i]['partno'])
   writexl.save('D:\deviceout\demo.xls')

def mergebycode():
   bk=xlrd.open_workbook('D:\\deviceout\\deviceout.xls', formatting_info=True)
   sh=bk.sheet_by_index(1)
   nrows=sh.nrows
   ncols=sh.ncols
   list=[]
   app={}
   for rownum in range(4,nrows):
      rowvalue=sh.row_values(rownum)
      app['id']=rowvalue[4]                ##总序號
      app['groupid']=rowvalue[5]           ##分单序号  
      app['partno']=rowvalue[21]           ##料号  
      app['origcountry']=rowvalue[19]      ##原产国  
      app['produceyear']=rowvalue[20]      ##产生年代  
      app['realmodel']=rowvalue[18]        ##实际品牌型号  
      app['realprice']=rowvalue[42]        ##实际出口价格  
      app['mergeid']=0                     ##归项  
      list.append(app.copy()) 

   for i in range(len(list)):
      compvalue=list[i]     
      if compvalue['mergeid']==0:
         compvalue['mergeid']=1
      for j in range(i+1,len(list)):
         ##录单版要点:EXCEL中的分单序号必须从小到大依次排序
         if compvalue['groupid']!=list[j]['groupid']:
            break
         elif compvalue['groupid']==list[j]['groupid']:
            if compvalue['partno']==list[j]['partno']:
               if compvalue['origcountry']==list[j]['origcountry'] and\
                  compvalue['produceyear']==list[j]['produceyear'] and \
                  compvalue['realmodel']==list[j]['realmodel'] and \
                  compvalue['realprice']==list[j]['realprice']:
                  list[j]['mergeid']=compvalue['mergeid']
               else:
                  if compvalue['mergeid']==20 :continue
                  if list[j]['mergeid']==compvalue['mergeid'] or list[j]['mergeid']==0:   
                     list[j]['mergeid']=compvalue['mergeid']+1    
            else:
               if compvalue['mergeid']==20 :continue
               if list[j]['mergeid']==compvalue['mergeid'] or list[j]['mergeid']==0: 
                  list[j]['mergeid']=compvalue['mergeid']+1
   '''        
            if compvalue['id']==121:
               print list[j]['id'],list[j]['groupid'],list[j]['partno'],list[j]['produceyear'] ,list[j]['mergeid'] ,list[j]['realmodel'] ,list[j]['origcountry'],  list[j]['realprice']    
              
   
   
   for ibk in list:            
      if  ibk['groupid'] in[25]:
         print ibk['id'],ibk['groupid'],ibk['partno'],ibk['produceyear'] ,ibk['mergeid'] ,ibk['realmodel'] ,ibk['origcountry'],  ibk['realprice']      
  
   
  
   wb=copy(bk)
   ws=wb.get_sheet(1)
   #styleBoldRed =xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: color-index red, bold on')
   styleBoldRed =xlwt.easyxf('font: color-index red, bold on')
   ws.write(2,ncols,u'归项')
   for ibk in range(len(list)):
      ws.write(ibk+4,ncols,list[ibk]['mergeid']) 
   wb.save('D:\deviceout\deviceout.xls')        
  
   '''
   writexl=xlwt.Workbook()
   sheet1=writexl.add_sheet(u'sheet1')
   sheet1.write(0,0,u'总序号')
   sheet1.write(0,1,u'归项')
   sheet1.write(0,2,u'料号') 
   for i in range(len(list)):
      sheet1.write(i+1,0,list[i]['id'])
      sheet1.write(i+1,1,list[i]['mergeid'])
      sheet1.write(i+1,2,list[i]['partno'])
   writexl.save('D:\deviceout\demo1.xls')


def relationCreate():
   bk=xlrd.open_workbook('D:\\deviceout\\relation_todo.xls')   
   sh=bk.sheet_by_index(1)
   nrows=sh.nrows
   ncols=sh.ncols
   list=[]
   app={}
   for rownum in range(4,nrows):
      rowvalue=sh.row_values(rownum)
      app['ybno']=rowvalue[0]           ##预报单号
      app['item']=rowvalue[1]           ##申报项次
      app['originno']=rowvalue[36]      ##原进口报关单号
      app['origoutno']=rowvalue[34]     ##原出口报关单号
      app['applyname']=rowvalue[102]    ##供应商名称
      app['applycode']=rowvalue[89]     ##供应商代码
      app['origapplydate']=rowvalue[85] ##原申报日期
      app['origitem']=rowvalue[69]      ##原申报项次
      app['origg_no']=rowvalue[73]      ##原备案序号
      app['origamount']=rowvalue[90]    ##原数量 
      app['origprice']=rowvalue[78]     ##原申报单价
      app['origcurrency']=rowvalue[79]  ##原申报币别
      app['origpassdate']=rowvalue[86]  ##原放行日期
      app['exchangerate']=rowvalue[87]  ##汇率
      app['price']=rowvalue[44]         ##申报单价
      app['outamount']=rowvalue[24]     ##出区数量
      list.append(app.copy()) 

   list=sorted(list,key=lambda app:(app['ybno'],app['item'])) #按预报单号和申报项次进行排序
   diclist =[]
   
   for i in range(len(list)):
      flag=False
      dic=list[i]
      for x in diclist:
         if x['ybno']==dic['ybno'] and x['item']==dic['item'] \
            and dic['originno']==x['originno'] and dic['origitem']==x['origitem'] :
            flag=True
            break             
      if flag:continue        
      for j in range(i+1,len(list)):
         if dic['ybno']==list[j]['ybno'] and dic['item']==list[j]['item'] \
            and dic['originno']==list[j]['originno'] and  dic['origitem']==list[j]['origitem']: 
            dic['outamount']+=list[j]['outamount']            
         else:
           diclist.append(dic)
           break 
   '''        
   for y in diclist:
      if y['ybno']=='1HX-7B0028':
         print y['ybno'],y['item'],y['originno'],y['origitem'],y['outamount']               
   '''        
   writexl=xlwt.Workbook()
   style =xlwt.easyxf('font: name SimSun,color-index red, bold on;align: wrap on, vert centre, horiz center')

   sheet1=writexl.add_sheet(u'sheet1')
   sheet1.write_merge(0,0,0,15,u'报关单对照关系表',style)
   sheet1.write(1,0,u'预报单号')
   sheet1.write(1,1,u'申报项次')
   sheet1.write(1,2,u'原进口报关单号')
   sheet1.write(1,3,u'原出口报关单号')   
   sheet1.write(1,4,u'供应商名称')   
   sheet1.write(1,5,u'供应商代码')          
   sheet1.write(1,6,u'原申报日期')   
   sheet1.write(1,7,u'原申报项次')   
   sheet1.write(1,8,u'原备案序号')   
   sheet1.write(1,9,u'原数量')   
   sheet1.write(1,10,u'原申报单价')
   sheet1.write(1,11,u'原申报币别')  
   sheet1.write(1,12,u'原放行日期')  
   sheet1.write(1,13,u'汇率')  
   sheet1.write(1,14,u'申报单价')  
   sheet1.write(1,15,u'出区数量')     
   for i in range(len(diclist)):
      sheet1.write(i+2,0,diclist[i]['ybno'])
      sheet1.write(i+2,1,diclist[i]['item'])
      sheet1.write(i+2,2,diclist[i]['originno'])
      sheet1.write(i+2,3,diclist[i]['origoutno'])
      sheet1.write(i+2,4,diclist[i]['applyname'])
      sheet1.write(i+2,5,diclist[i]['applycode'])
      sheet1.write(i+2,6,diclist[i]['origapplydate'])
      sheet1.write(i+2,7,diclist[i]['origitem'])
      sheet1.write(i+2,8,diclist[i]['origg_no'])
      sheet1.write(i+2,9,diclist[i]['origamount'])
      sheet1.write(i+2,10,diclist[i]['origprice'])
      sheet1.write(i+2,11,diclist[i]['origcurrency'])
      sheet1.write(i+2,12,diclist[i]['origpassdate'])
      sheet1.write(i+2,13,diclist[i]['exchangerate'])
      sheet1.write(i+2,14,diclist[i]['price'])
      sheet1.write(i+2,15,diclist[i]['outamount'])
   writexl.save('D:\deviceout\\relation.xls')
   

if __name__=='__main__':
   if os.path.exists('D:\\deviceout\\pre_deviceout.xls'):
      groupbycode()
   if os.path.exists('D:\\deviceout\\deviceout.xls'):
      mergebycode()   
   if os.path.exists('D:\\deviceout\\relation_todo.xls'):
      relationCreate()   