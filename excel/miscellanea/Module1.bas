Attribute VB_Name = "Module1"
Dim compGroup As String
Dim choiceSht As Worksheet, templateSht As Worksheet

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
    Dim desktopPath As String
    
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

Sub createChoises()
    'create comboboxes from choises text file
    Dim flw As New FileWorker, clw As New CellWorker
    Dim i As Integer
    Dim tmpCol As New Collection
    Dim inFile As String
    Dim myFile As String
    Dim tmpCell As Range
    Dim idCol As Integer
    Dim compEntryCol As Integer
    
    Dim excelName As String, nameRange As String
    Dim nameMustBeStarted As Boolean
    Dim tmpArr As Variant, tmpStr As String

    
    
    Set choiceSht = ActiveWorkbook.Sheets("choices") '@todo replace by actual file name
    Set templateSht = ActiveWorkbook.Sheets("template")
    choiceSht.Activate
    Set tmpCell = Range("A1")
    desktopPath = Environ("USERPROFILE")
    
    desktopPath = desktopPath & "\Desktop\"
    myFile = desktopPath & "choices.txt"
    
    Set tmpCol = flw.readLinesFromTxt(myFile)
    
    For Each LineItem In tmpCol
        tmpCell.value = LineItem
        Set tmpCell = clw.move_down(tmpCell)
    Next LineItem
    
    templateSht.Select
    Set tmpCell = Range("A3")
    Do While tmpCell.value <> ""
        tmpArr = Split(Cells(tmpCell.Row, 3).value, " ")
        For i = 0 To UBound(tmpArr)
            tmpStr = tmpStr & Left(LCase(tmpArr(i)), 1)
        Next i
        
        nameRange = computeRange(tmpCell.value)
        
        excelName = tmpStr & tmpCell.value
        ActiveWorkbook.Names.Add Name:=excelName, RefersToR1C1:="=choices!" & nameRange
        
        With Cells(tmpCell.Row, 4).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & excelName
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        
        Set tmpCell = clw.move_down(tmpCell)
        tmpStr = ""
        excelName = ""
    Loop

End Sub

Function computeRange(inInt As String) As String
    
    Dim intCell As Range
    Dim upCell As Range, downCell As Range
    Dim tmpCell As Range
    Dim clw As New CellWorker
    Dim tmpStr As String
    
    choiceSht.Activate
    Columns("A:A").Select
    Range("A1").Activate
    Set intCell = Selection.Find(What:=inInt, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
    Set upCell = Cells(intCell.Row + 1, intCell.Column)
    upCell.Select
    Set tmpCell = upCell
    Do While Not IsNumeric(tmpCell.value) And tmpCell.value <> ""
        Set tmpCell = clw.move_down(tmpCell)
    Loop
    Set downCell = Cells(tmpCell.Row - 1, tmpCell.Column)
    
    tmpStr = upCell.Address(ReferenceStyle:=xlR1C1) & ":" & downCell.Address(ReferenceStyle:=xlR1C1)
    
   templateSht.Activate
   
   computeRange = tmpStr
End Function
