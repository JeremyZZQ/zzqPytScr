Attribute VB_Name = "Module11"
Sub hz()
    ThisWorkbook.Sheets("利润预测表").Activate
    Call MacroLRYCB
    
    ThisWorkbook.Sheets("营收").Activate
    Call MacroYYSR
    
    ThisWorkbook.Sheets("营成").Activate
    Call MacroYYCB
    
    ThisWorkbook.Sheets("销费").Activate
    Call MacroXSFY
    
    ThisWorkbook.Sheets("管费").Activate
    Call MacroGLFY
    
    ThisWorkbook.Sheets("财费").Activate
    Call MacroCWFY
    
    ThisWorkbook.Sheets("资减损").Activate
    Call MacroZJS
    
    ThisWorkbook.Sheets("信减损").Activate
    Call MacroXJS
    
    ThisWorkbook.Sheets("三项收益").Activate
    Call MacroSXSY
    
    ThisWorkbook.Sheets("营业外收支").Activate
    Call MacroYYWSZ
    
    ThisWorkbook.Sheets("所得税费用").Activate
    Call MacroSDSFY
    
    ThisWorkbook.Sheets("少数股东损益").Activate
    Call MacroSSGDSY
    
End Sub

Sub MacroLRYCB()
'
' 利润预测表
'

'
    Range("C6:K30").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]利润预测表'!R6C3:R30C11", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]利润预测表'!R6C3:R30C11"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroYYSR()
'
' 营业收入
'

'
    Range("B2:H10").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]营收'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]营收'!R2C2:R10C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroYYCB()
Attribute MacroYYCB.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 营业成本
'

'
    Range("B2:H10").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]营成'!R2C2:R10C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]营成'!R2C2:R10C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroXSFY()
'
' 销售费用
'

'
    Range("B2:H18").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]销费'!R2C2:R18C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]销费'!R2C2:R18C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroGLFY()
'
' 管理费用
'

'
    Range("B2:H25").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]管费'!R2C2:R25C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]管费'!R2C2:R25C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub
Sub MacroCWFY()
'
' 财务费用
'

'
    Range("B2:H23").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]财费'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]财费'!R2C2:R23C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroZJS()
'
' 资产减值损失
'

'
    Range("B2:H8").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]资减损'!R2C2:R8C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]资减损'!R2C2:R8C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroXJS()
'
' 信用减值损失
'

'
    Range("B2:H12").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]信减损'!R2C2:R12C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]信减损'!R2C2:R12C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub


Sub MacroSXSY()
'
' 三项收益
'

'
    Range("B2:H23").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]三项收益'!R2C2:R23C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]三项收益'!R2C2:R23C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub



Sub MacroYYWSZ()
'
' 营业外收支
'

'
    Range("B2:H20").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]营业外收支'!R2C2:R20C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]营业外收支'!R2C2:R20C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroSDSFY()
'
' 所得税费用
'

'
    Range("B2:H7").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]所得税费用'!R2C2:R7C8", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]所得税费用'!R2C2:R7C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroSSGDSY()
'
' 少数股东损益
'

'
    Range("B2:I8").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛CONGOBEST.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛埃塞.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛安哥拉.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛本部.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛迪拜.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛刚果布.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛吉布提.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛加纳.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛喀麦隆贸易.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛马里.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛南苏丹.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛尼日利亚pvc.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛坦桑.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乌干达.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛伊拉克.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛赞比亚.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_汉盛乍得项目组.xlsx]少数股东损益'!R2C2:R8C9", _
        "'F:\预算\利润表预测2019年8月\data\[利润预测表_京胜合并.xlsx]少数股东损益'!R2C2:R8C9"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub
