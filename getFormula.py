#!python3

'''
读取利润预测表模板的公式

'''
import openpyxl
import re
import pprint
fmlRegex=re.compile(r'^={1}\D+.*',re.IGNORECASE)

#todo:tkinter选择文件，检查sheet名称是否被修改

wb=openpyxl.load_workbook(r'F:\预算\利润表预测2019年8月\data\利润预测表new.xlsx')

rg={'利润预测表':'c6:k30',
    '营收':'b2:h11',
    '营成':'b2:h11',
    '销费':'b2:h19',
    '管费':'b2:h26',
    '财费':'b2:h24',
    '资减损':'b2:h9',
    '信减损':'b2:h13',
    '三项收益':'b2:h24',
    '营业外收支':'b2:h21',
    '所得税费用':'b2:h7',
    '少数股东损益':'b2:i8'
    }
fm={}
rgfm={}

for k in rg.keys():
    #if k!='利润预测表':
    #   break
    
    rgfm.setdefault(k,None)
    sheet=wb[k]
    
    for rowOfCellObjects in sheet[rg[k]]:
        for cellObj in rowOfCellObjects:
            isfml=fmlRegex.search(str(cellObj.value))
            if isfml!=None and len(isfml.group())>0:
                fm[cellObj.coordinate]=cellObj.value
                
        
    rgfm[k]=fm


fmlFile=open('newFormula.txt','w')
fmlFile.write(pprint.pformat(rgfm))
fmlFile.close