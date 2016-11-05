Attribute VB_Name = "Module1"
Sub 상대참조_복사_테스트()
Attribute 상대참조_복사_테스트.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' 상대참조_복사_테스트 매크로
'
' 바로 가기 키: Ctrl+e
'
    Dim headerRange, studentRange, resultRange As Range
    Dim startCellAddress, searchString As String
    startCellAddress = "E1"
    searchString = "누적"
    
    With ActiveSheet.Range(startCellAddress + ":AZ3")
        Set searchResult = .Find(searchString, LookIn:=xlValues)
        
        Set headerRange = Range(startCellAddress + ":" + Cells(searchResult.Row + 2, searchResult.Column).Address)
        Set studentRange = ActiveCell.Range("A1:" + Cells(1, searchResult.Column - Range(startCellAddress).Column + 1).Address)
        
        Set resultRange = union(headerRange, studentRange)
        resultRange.Copy
    End With
    
End Sub
Sub 학생숨김_후_이동()
Attribute 학생숨김_후_이동.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' 학생숨김_후_이동 매크로
'
' 바로 가기 키: Ctrl+r
'
        
    ActiveCell.Select
    Selection.EntireRow.Hidden = True
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

Sub 토일_복사_매크로()
'
' 토일_복사_매크로 매크로
'
' 바로 가기 키: Ctrl+t
'
    'Dim clipboard As MSForms.DataObject
    'Set clipboard = New MSForms.DataObject
    Dim temp As Range
    Dim day, prefix, postfix, resultString, todayString As String
    
    Set temp = Range("BB1")
    prefix = "안녕하세요? 대찬학원 임재상 선생님의 조교입니다. 한영외고1 내신 1차시 수업("
    postfix = ") 관리표를 첨부하여 보냅니다."
    
    day = ActiveCell.Previous.Value
    If day = "토" Then
        todayString = Format(Date - 1, "m월 d일")
        resultString = prefix + todayString + "/" + day + postfix
    Else
        todayString = Format(Date, "m월 d일")
        resultString = prefix + todayString + "/" + day + postfix
    End If
    
    temp.Value = resultString
    temp.Copy
    
End Sub
Sub 엄마_복사_매크로()
'
' 엄마_복사_매크로 매크로
'
' 바로 가기 키: Ctrl+d
'
    
    ActiveCell.Offset(0, -2).Copy
    
    
End Sub

