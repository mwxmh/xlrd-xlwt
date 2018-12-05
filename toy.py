#coding:utf-8
from xlutils.copy import copy 
import xlrd,xlwt,sys,os,datetime
reload(sys)
sys.setdefaultencoding('utf-8')
def main():
    bk=xlrd.open_workbook('D:\\deviceout\\detail.xlsx')    
    sh=bk.sheet_by_index(1)
    nrows=sh.nrows
    ncols=sh.ncols
    list=[]
    app={}    
    for rownum in range(0,nrows):
        rowvalue=sh.row_values(rownum)
        app['gonglingNo']=rowvalue[1].strip()           #工令号码
        app['inPutDate']=rowvalue[2].strip()            #输入日期
        app['materialno']=rowvalue[3].strip()           #成品料号 
        app['CustomerPN']=rowvalue[4].strip()            #Customer p/n
        app['materialNam']=rowvalue[5].strip()            #品名规格
        app['produceNum']=rowvalue[6]                   #生产数量
        app['HasSentNum']=rowvalue[7]                   #已发套数    
        app['FinishNum']=rowvalue[8]                   #完工数量
        app['NofinishNum']=rowvalue[9]                   #未完工数量
        app['NeedTime']=rowvalue[10]                    #投入工时
        app['FreeDays']=rowvalue[11]                   #休息天数
        app['IsDelay']=rowvalue[12].strip()              #是否超期
        app['StartDay']=rowvalue[13].strip()              #开工日期
        app['preFinshday']=rowvalue[14].strip()              #预完工日
        app['Stat']=rowvalue[15]                  #状态
        app['groupClass']=rowvalue[16]             #组别
        app['S/O']=rowvalue[17].strip()              #S/O
        app['Module']=rowvalue[18].strip()              #Module
        app['StandardCost']=rowvalue[19]              #标准成本
        app['NofinishCost']=rowvalue[20]              #未结案金额KRMB
        app['TargetDate']=rowvalue[21].strip()              #TargetDate
        app['STDDays']=rowvalue[22]                #STDDays
        app['AgingDays']=rowvalue[23]                 #Aging Days
        app['Agingtype']=rowvalue[24].strip()                 #Aging type
        app['ID']=rowvalue[25].strip()                        #ID
        app['WO']=rowvalue[26].strip()                 #WO
        app['Linetype']=rowvalue[27].strip()                 #主線or前加工
        app['Level']=rowvalue[28].strip()                 #阶段
        app['WOType']=rowvalue[29].strip()                 #WOType
        app['isFinish']=rowvalue[30].strip()                 #可結案否
        app['analyField']=rowvalue[31].strip()                 #分析栏位
        app['FeeCode']=rowvalue[32].strip()                 #费用代码
        list.append(app.copy())
    print len(list)    

    tupl_remove=('CR26','CR2B','CR2E','CR2H','CR2I','CR2J','CR2N','CR2P','CR2S','CR2U','CR2W','CW29',\
                 'CW2K','CW2M','CW2T','CW2U','CW2V','CW2Z','DR22','DW22')                
    tupl_7=('CW20','CW24','CW27','CW28','CW2A','CW2C','CW2D','CW2F','CW2P','CW33','CW34','CW3A','CW3U',\
            'CW3Y')
    for i in range(len(list)-1, -1, -1): #倒序
        if list[i]['ID'] in tupl_remove:
            list.pop(i)
            #list.remove(list[i]) 
        if list[i]['ID'] in tupl_7:
            datestr=list[i]['StartDay'] #开工日期
            date=datetime.datetime.strptime(datestr,'%Y-%m-%d')
            today=datetime.datetime.now() #今天
            n_days=today-date
            if n_days.days<8:
                list.pop(i)

    
    writexl=xlwt.Workbook() 
    sheet1=writexl.add_sheet(u'sheet1')
    for i in range(len(list)): 
        sheet1.write(i,0,list[i]['gonglingNo'])
        sheet1.write(i,1,list[i]['inPutDate'])
        sheet1.write(i,2,list[i]['materialno'])
        sheet1.write(i,3,list[i]['CustomerPN'])
        sheet1.write(i,4,list[i]['materialNam'])
        sheet1.write(i,5,list[i]['produceNum'])
        sheet1.write(i,6,list[i]['HasSentNum'])
        sheet1.write(i,7,list[i]['FinishNum'])          
        sheet1.write(i,8,list[i]['NofinishNum'])
        sheet1.write(i,9,list[i]['NeedTime'])
        sheet1.write(i,10,list[i]['FreeDays'])
        sheet1.write(i,11,list[i]['IsDelay'])
        sheet1.write(i,12,list[i]['StartDay'])
        sheet1.write(i,13,list[i]['preFinshday'])
        sheet1.write(i,14,list[i]['Stat'])
        sheet1.write(i,15,list[i]['groupClass'])
        sheet1.write(i,16,list[i]['S/O'])
        sheet1.write(i,17,list[i]['Module'])
        sheet1.write(i,18,list[i]['StandardCost'])
        sheet1.write(i,19,list[i]['NofinishCost'])
        sheet1.write(i,20,list[i]['TargetDate'])
        sheet1.write(i,21,list[i]['STDDays'])
        sheet1.write(i,22,list[i]['AgingDays'])
        sheet1.write(i,23,list[i]['Agingtype'])
        sheet1.write(i,24,list[i]['ID'])
        sheet1.write(i,25,list[i]['WO'])
        sheet1.write(i,26,list[i]['Linetype'])
        sheet1.write(i,27,list[i]['Level'])
        sheet1.write(i,28,list[i]['WOType'])
        sheet1.write(i,29,list[i]['isFinish'])
        sheet1.write(i,30,list[i]['analyField'])
        sheet1.write(i,31,list[i]['FeeCode'])
    writexl.save('D:\deviceout\\detailPro.xls')    
    



        
if __name__ == '__main__':
    main()