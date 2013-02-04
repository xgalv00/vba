Attribute VB_Name = "Module1"
Dim compGroup As String

Sub collectComs()
    Dim finalMsg As String
    Dim tmpCell As Range
    Dim clw As New CellWorker
    Dim desktopPath As String
    Dim myFile As String
    
    ActiveWorkbook.Sheets("template").Activate
    Set tmpCell = Cells(3, 3)
    
    Do While tmpCell.value <> ""
        finalMsg = finalMsg & createMsg(tmpCell.Row)
        Set tmpCell = clw.move_down(tmpCell)
    Loop
    
    desktopPath = Environ("USERPROFILE")
    
    desktopPath = desktopPath & "\Desktop\"
    myFile = desktopPath & "selfassessment_result.txt"
    
    Open myFile For Output As 1

    Print #1, finalMsg
    
    Close #1
    
    
End Sub

Function createMsg(rowToPrcs As Integer) As String

    Dim tmpCompGroup As String, compGroupCol As Integer
    Dim compEntry As String, compEntryCol As Integer
    Dim example As String, exampleCol As Integer
    Dim compGroupChanged As Boolean
    Dim tmpCell As Range, msg As String
    
    compGroupCol = 2
    compEntryCol = 3
    exampleCol = 5
    
    tmpCompGroup = Cells(rowToPrcs, compGroupCol).value
    compEntry = Cells(rowToPrcs, compEntryCol).value
    example = Cells(rowToPrcs, exampleCol).value
    
    If Not tmpCompGroup = compGroup Then
        compGroup = tmpCompGroup
        compGroupChanged = True
    End If
    If compGroupChanged Then
        msg = compGroup & vbCrLf & "Компетенция: " & compEntry & vbCrLf & "Комментарий: " & example & vbCrLf
    Else
        msg = "Компетенция " & compEntry & vbCrLf & "Комментарий: " & example & vbCrLf
    End If
    
    createMsg = msg

End Function
