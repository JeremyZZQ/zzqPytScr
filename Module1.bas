Attribute VB_Name = "Module11"
Sub hz()
    ThisWorkbook.Sheets("����Ԥ���").Activate
    Call MacroLRYCB
    
    ThisWorkbook.Sheets("Ӫ��").Activate
    Call MacroYYSR
    
    ThisWorkbook.Sheets("Ӫ��").Activate
    Call MacroYYCB
    
    ThisWorkbook.Sheets("����").Activate
    Call MacroXSFY
    
    ThisWorkbook.Sheets("�ܷ�").Activate
    Call MacroGLFY
    
    ThisWorkbook.Sheets("�Ʒ�").Activate
    Call MacroCWFY
    
    ThisWorkbook.Sheets("�ʼ���").Activate
    Call MacroZJS
    
    ThisWorkbook.Sheets("�ż���").Activate
    Call MacroXJS
    
    ThisWorkbook.Sheets("��������").Activate
    Call MacroSXSY
    
    ThisWorkbook.Sheets("Ӫҵ����֧").Activate
    Call MacroYYWSZ
    
    ThisWorkbook.Sheets("����˰����").Activate
    Call MacroSDSFY
    
    ThisWorkbook.Sheets("�����ɶ�����").Activate
    Call MacroSSGDSY
    
End Sub

Sub MacroLRYCB()
'
' ����Ԥ���
'

'
    Range("C6:K30").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]����Ԥ���'!R6C3:R30C11", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]����Ԥ���'!R6C3:R30C11"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroYYSR()
'
' Ӫҵ����
'

'
    Range("B2:H10").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]Ӫ��'!R2C2:R10C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroYYCB()
Attribute MacroYYCB.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Ӫҵ�ɱ�
'

'
    Range("B2:H10").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]Ӫ��'!R2C2:R10C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]Ӫ��'!R2C2:R10C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroXSFY()
'
' ���۷���
'

'
    Range("B2:H18").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]����'!R2C2:R18C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]����'!R2C2:R18C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroGLFY()
'
' ��������
'

'
    Range("B2:H25").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]�ܷ�'!R2C2:R25C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]�ܷ�'!R2C2:R25C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub
Sub MacroCWFY()
'
' �������
'

'
    Range("B2:H23").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]�Ʒ�'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]�Ʒ�'!R2C2:R23C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroZJS()
'
' �ʲ���ֵ��ʧ
'

'
    Range("B2:H8").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]�ʼ���'!R2C2:R8C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]�ʼ���'!R2C2:R8C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroXJS()
'
' ���ü�ֵ��ʧ
'

'
    Range("B2:H12").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]�ż���'!R2C2:R12C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]�ż���'!R2C2:R12C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub


Sub MacroSXSY()
'
' ��������
'

'
    Range("B2:H23").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]��������'!R2C2:R23C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]��������'!R2C2:R23C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub



Sub MacroYYWSZ()
'
' Ӫҵ����֧
'

'
    Range("B2:H20").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]Ӫҵ����֧'!R2C2:R20C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]Ӫҵ����֧'!R2C2:R20C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroSDSFY()
'
' ����˰����
'

'
    Range("B2:H7").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]����˰����'!R2C2:R7C8", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]����˰����'!R2C2:R7C8"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub

Sub MacroSSGDSY()
'
' �����ɶ�����
'

'
    Range("B2:I8").Select
    Selection.Consolidate Sources:=Array( _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢCONGOBEST.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ϰ�.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�չ���.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����¡ó��.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ����.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ���յ�.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ��������pvc.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ̹ɣ.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ڸɴ�.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ������.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢ�ޱ���.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʢէ����Ŀ��.xlsx]�����ɶ�����'!R2C2:R8C9", _
        "'F:\Ԥ��\�����Ԥ��2019��8��\data\[����Ԥ���_��ʤ�ϲ�.xlsx]�����ɶ�����'!R2C2:R8C9"), Function:=xlSum _
        , TopRow:=False, LeftColumn:=False, CreateLinks:=True
End Sub