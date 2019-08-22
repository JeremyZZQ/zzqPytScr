#!python3

'''
对指定区域空白单元格批量填零，以便合并计算

'''

import openpyxl
subsidiary = ('汉盛本部',
              '汉盛赞比亚',
              '汉盛安哥拉',
              '汉盛CONGOBEST',
              '汉盛尼日利亚',
              '汉盛埃塞',
              '京胜合并',
              '汉盛马里',
              '汉盛伊拉克',
              '汉盛吉布提',
              '汉盛喀麦隆贸易',
              '汉盛乍得项目组',
              '汉盛刚果布',
              '汉盛坦桑',
              '汉盛乌干达',
              '汉盛迪拜',
              '汉盛尼日利亚pvc',
              '汉盛南苏丹',
              '汉盛加纳')

sheetsAndRange = {'利润预测表': 'c6:k30',
                  '营收': 'b2:h10',
                  '营成': 'b2:h10',
                  '销费': 'b2:h18',
                  '管费': 'b2:h25',
                  '财费': 'b2:h23',
                  '资减损': 'b2:h8',
                  '信减损': 'b2:h12',
                  '三项收益': 'b2:h23',
                  '营业外收支': 'b2:h20',
                  '所得税费用': 'b2:h7',
                  '少数股东损益': 'b2:i8'
                  }

fileNameEx = 'F:\预算\利润表预测2019年8月\data\利润预测表_'
newFileNameEx='F:\预算\利润表预测2019年8月\data\修改\利润预测表_'

for sub in range(19):

    fileName = fileNameEx+subsidiary[sub]+'.xlsx'
    newFileName=newFileNameEx+subsidiary[sub]+'.xlsx'
    wb = openpyxl.load_workbook(fileName)

    for sht in sheetsAndRange.keys():
        sheet=wb[sht]
        for rowOfCellObjects in sheet[sheetsAndRange[sht]]:
            for cellObj in rowOfCellObjects:
                if cellObj.value == None:
                    sheet[cellObj.coordinate] = 0
                    print(cellObj.coordinate,sheet[cellObj.coordinate].value)
        wb.save(newFileName)
