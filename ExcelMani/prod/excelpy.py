'''
Created on Mar 9, 2014

@author: Yusheng Ding

.ExcelMani.excelOpen
'''

import pyExcelerator as pyEx

import sys


def readExcel(filePath):   #read excel --> list [(row_idx, col_idx), value]
    data=[]
    try:    
        #  
        for sheet_name, values in pyEx.parse_xls(filePath, 'cp1251'): # parse_xls(arg) -- default encoding
            print 'Sheet = "%s"' % sheet_name.encode('cp866', 'backslashreplace')
            print '----------------'
            for row_idx, col_idx in sorted(values.keys()):
                v = values[(row_idx, col_idx)]
                if isinstance(v, unicode):
                    v = v.encode('cp866', 'backslashreplace')
                print '(%d, %d) =' % (row_idx, col_idx), v
                data.append([(row_idx, col_idx), v])
            print '----------------'
            #print 'file read successful'
            return data
    except:
        print 'file read Error'
        print "Unexpected error:", sys.exc_info()
        return -1
#


def writExcel(listData,fileName,sheetName):  #write listData -->Excel File   list [(row_idx, col_idx), value]
    try:
        
        print 'Writing file, please wait......'
        
        w = pyEx.Workbook()
        ws = w.add_sheet(sheetName)
        for location in listData:
            ws.write(location[0][0],location[0][1],location[1])       
        w.save(fileName)
        print  'file write successful-->','filePath: ', fileName
        return 0
    except:
        print 'file write Error'
        print "Unexpected error:", sys.exc_info()
        return -1
    
    
    
def anotherListData(listData,numColumn=4):  #change dataList  list [(row_idx, col_idx), value] into  ['value','value','value','value','value']
    newList=[]
    tempList=[]
    try:
        for data in listData:
            if(data[0][1]!=numColumn):
                tempList.append(data[1]) 
                #print tempList          
            else:
                tempList.append(data[1]) 
                newList.append(tempList)
                tempList=[]
                    
        #print 'another List successful'
        return newList
    except:
        print 'another List Error'
        print "Unexpected error:", sys.exc_info()
        return -1
   
        
        
        
def changeData(listData):     
    try:
        newList=[]
        count=0
        countTotal=len(listData)
        for value in listData:
            if(count==0 or count==countTotal-1):
                newList.append(value)
            else:
                valueTemp=value
                
                i=listData[count+1][3]-valueTemp[3]
                while(i>1):       
                    newList.append([0,0,0,valueTemp[3],0])    # fill in data
                    valueTemp[3]=valueTemp[3]+1
                    i=i-1
                    
                else:
                    newList.append(value)
            count = count + 1
          

        return newList
    except:
        print 'Error change data'
        print "Unexpected error:", sys.exc_info()
        return -1



def listToExcelList(listData,numColumn=5):  #change dataList  ['value','value','value','value','value'] into list [(row_idx, col_idx), value]
    newList=[]
    locationList=[]
    try:
        for numRaw in range(0,len(listData)):
            for nC in range(0,numColumn):
                locationList.append((numRaw,nC))

        #print locationList
        i=0
        for data in listData:
            for value in data:
                if(i<len(locationList)):
                    newList.append([locationList[i],value])
                    i=i+1  
        return newList
                
                
                
    except:
        print 'listToExcelList Error'
        print "Unexpected error:", sys.exc_info()
        return -1
    
  



#split dataList
def splitDataList(listData):
    try:
        list1=[]
        list2=[]
        newlist=[]
        numLen=len(listData)
        index1=0
        index2=750000
        while index1 < 750000:
            list1.append(listData[index1])
            index1=index1+1
        while index2 < numLen:
            list2.append(listData[index2])
            index2=index2+1
            
        newlist.append([list1])
        newlist.append([list2])  
        return newlist
    except:
        print 'splitDataList Error'
        print "Unexpected error:", sys.exc_info()
        return -1
    
def split_seq(seq, p):
    newseq = []
    n = len(seq) / p    # min items per subsequence
    r = len(seq) % p    # remaindered items
    b,e = 0, n + min(1, r)  # first split
    for i in range(p):
        newseq.append(seq[b:e])
        r = max(0, r-1)  # use up remainders
        b,e = e, e + n + min(1, r)  # min(1,r) is always 0 or 1
    
    return newseq    

'''
'''
    

# filePath='../excelFile/'
# oldfileName='SEC_PART1.xls'
# newfileName1='generatePart1.xls'
# newSheetName='sheet01'
# 
# 
# oldfilePath=filePath+oldfileName
# 
# newfilePath1=filePath+newfileName1
# 
# exdata=readExcel(oldfilePath)
# 
# data2 = anotherListData(exdata,4)
# 
# data3 = changeData(data2)
# 
# dataFinal1= listToExcelList(data3)
# 
# writExcel(dataFinal1,newfilePath1,newSheetName)
# ep.writExcel(dataFinal2,newfilePath2,newSheetName)

