'''
Created on Mar 10, 2014

@author: Yusheng Ding

.ExcelMani.prod.main
'''
from ExcelMani.prod import excelpy as ep


filePath='../excelFile/'
newFilePath=''
oldfileName='cdctwitter_FB_Sec_Groups_Big.xls'
#newfileName1='generatePart2.xls'
#newSheetName='sheet02'

oldfilePath=filePath+oldfileName


# 
# newfilePath=filePath+newfileName1
# 
# exdata=ep.readExcel(oldfilePath)
# 
# data2 = ep.anotherListData(exdata,4)
# 
# data3 = ep.changeData(data2)
# 
# dataFinal=ep.listToExcelList(data3)
# 
# ep.writExcel(dataFinal,newfilePath,newSheetName)
# ep.writExcel(dataFinal2,newfilePath2,newSheetName)

newfileList=[]

dataFinal=[]
numofFile=45
for i in range(1,numofFile):
    fileName='generatePart'+str(i)+'.xls'
    filePath=filePath+newFilePath
    newfileList.append(fileName)

exdata=ep.readExcel(oldfilePath)
data2=ep.anotherListData(exdata,4)
data3=ep.changeData(data2)
data4=ep.split_seq(data3,numofFile)

i=0
for data in data4:
    ep.writExcel(ep.listToExcelList(data), newfileList[i],'sheet1')
    i=i+1


    

