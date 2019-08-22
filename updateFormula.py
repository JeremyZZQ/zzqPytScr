#!python3

'''
更新利润预测表的公式

'''

import openpyxl
import os

newFml = {'利润预测表': {'E10': '=SUM(F10:K10)',
                    'E11': '=SUM(F11:K11)',
                    'E12': '=SUM(F12:K12)',
                    'E13': '=SUM(F13:K13)',
                    'E14': '=SUM(F14:K14)',
                    'E15': '=SUM(F15:K15)',
                    'E16': '=SUM(F16:K16)',
                    'E17': '=SUM(F17:K17)',
                    'E18': '=SUM(F18:K18)',
                    'E19': '=SUM(F19:K19)',
                    'E20': '=SUM(F20:K20)',
                    'E21': '=SUM(F21:K21)',
                    'E22': '=SUM(F22:K22)',
                    'E23': '=SUM(F23:K23)',
                    'E24': '=SUM(F24:K24)',
                    'E25': '=SUM(F25:K25)',
                    'E26': '=SUM(F26:K26)',
                    'E27': '=SUM(F27:K27)',
                    'E28': '=SUM(F28:K28)',
                    'E29': '=SUM(F29:K29)',
                    'E30': '=SUM(F30:K30)',
                    'E7': '=SUM(F7:K7)',
                    'E8': '=SUM(F8:K8)',
                    'E9': '=SUM(F9:K9)'}}

wb=openpyxl.load_workbook(r'F:\预算\利润表预测2019年8月\data\利润预测表test.xlsx')

for sheetk in newFml.keys():
    updateFml=newFml[sheetk]
    updateSht=wb[sheetk]
    for cellk in updateFml.keys():
        updateSht[cellk]=updateFml(cellk)

