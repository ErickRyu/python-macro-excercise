Attribute VB_Name = "Module1"
Sub �������_����_�׽�Ʈ()
Attribute �������_����_�׽�Ʈ.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' �������_����_�׽�Ʈ ��ũ��
'
' �ٷ� ���� Ű: Ctrl+e
'
    Dim headerRange, studentRange, resultRange As Range
    Dim startCellAddress, searchString As String
    startCellAddress = "E1"
    searchString = "����"
    
    With ActiveSheet.Range(startCellAddress + ":AZ3")
        Set searchResult = .Find(searchString, LookIn:=xlValues)
        
        Set headerRange = Range(startCellAddress + ":" + Cells(searchResult.Row + 2, searchResult.Column).Address)
        Set studentRange = ActiveCell.Range("A1:" + Cells(1, searchResult.Column - Range(startCellAddress).Column + 1).Address)
        
        Set resultRange = union(headerRange, studentRange)
        resultRange.Copy
    End With
    
End Sub
Sub �л�����_��_�̵�()
Attribute �л�����_��_�̵�.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' �л�����_��_�̵� ��ũ��
'
' �ٷ� ���� Ű: Ctrl+r
'
        
    ActiveCell.Select
    Selection.EntireRow.Hidden = True
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

Sub ����_����_��ũ��()
'
' ����_����_��ũ�� ��ũ��
'
' �ٷ� ���� Ű: Ctrl+t
'
    'Dim clipboard As MSForms.DataObject
    'Set clipboard = New MSForms.DataObject
    Dim temp As Range
    Dim day, prefix, postfix, resultString, todayString As String
    
    Set temp = Range("BB1")
    prefix = "�ȳ��ϼ���? �����п� ����� �������� �����Դϴ�. �ѿ��ܰ�1 ���� 1���� ����("
    postfix = ") ����ǥ�� ÷���Ͽ� �����ϴ�."
    
    day = ActiveCell.Previous.Value
    If day = "��" Then
        todayString = Format(Date - 1, "m�� d��")
        resultString = prefix + todayString + "/" + day + postfix
    Else
        todayString = Format(Date, "m�� d��")
        resultString = prefix + todayString + "/" + day + postfix
    End If
    
    temp.Value = resultString
    temp.Copy
    
End Sub
Sub ����_����_��ũ��()
'
' ����_����_��ũ�� ��ũ��
'
' �ٷ� ���� Ű: Ctrl+d
'
    
    ActiveCell.Offset(0, -2).Copy
    
    
End Sub

