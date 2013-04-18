Attribute VB_Name = "Administration"
Dim relToRange As Range
Sub checkMineRange()
    Dim tmpSht As Worksheet
    Set tmpSht = Sheets("control_table_ÁÏÑÑ_ø")
    tmpSht.Visible = xlSheetVisible
    tmpSht.Select
    Set relToRange = Range("E11")
    Application.EnableEvents = False
    Call processMineRange(Range("A1"))
    Application.EnableEvents = True
End Sub

'check computed addresses
Private Sub processMineRange(upLeftCell As Range)

    'Returns collection of addresses that should be copied
    Dim selRange As Range
    Dim i As Integer
    Dim rangeAddr As String
 
    
    
    Call moveThroughRows(upLeftCell)
    
    Debug.Assert addrColl.Count > 1
    
    Call activateApprSht(ActiveSheet.Name)
    
    For Each addr In addrColl
        If selRange Is Nothing Then
            Set selRange = Range(addr)
        Else
            Set selRange = Application.Union(selRange, Range(addr))
        End If
    Next addr
    
    
    selRange.Select
    
End Sub


Private Sub activateApprSht(curShtName As String)

    'activates sheet appropriate to active control sheet
    
    Dim tmpStr As String
    
    tmpStr = Right(curShtName, Len(curShtName) - Len("control_table_"))
    
    ActiveWorkbook.Sheets(tmpStr).Activate
    

End Sub

'Function returnRangeAddr(tmpRange As Range) As String
'
'    returnRangeAddr = tmpRange.Address(False, False)
'
'End Function
''
'Function convertToR1C1(tmpRange As Range, relativeTo As Range) As String
'
'    Dim tmpString As String
'    '>>>convertToR1C1(Range("G17:H18"),Range("E9"))
'    'R[8]C[2]:R[9]C[3]
'    'Debug.Print ""
'    tmpString = tmpRange.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, relativeTo:=relativeTo)
'    convertToR1C1 = tmpString
'End Function
Sub processSelRow()
    Dim tmpStr As String
    Dim upLeftCell As Range
    Dim clw As New CellWorker
    
    Set relToRange = Range("E128")
    Set upLeftCell = Sheets("control_table_" & ActiveSheet.Name).Range("A1")
    For Each areaItem In Selection.Areas
        tmpStr = areaItem.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, relativeTo:=relToRange)
        upLeftCell.value = tmpStr
        tmpStr = ""
        Set upLeftCell = clw.move_right(upLeftCell, 2)
    Next areaItem

End Sub

Sub processSelCol()

    Dim tmpStr As String
    Dim upLeftCell As Range
    Dim clw As New CellWorker
    
    Set relToRange = Range("E128")
    Set upLeftCell = Sheets("control_table_" & ActiveSheet.Name).Range("A1")
    For Each areaItem In Selection.Areas
        tmpStr = areaItem.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, relativeTo:=relToRange)
        upLeftCell.value = tmpStr
        tmpStr = ""
        Set upLeftCell = clw.move_down(upLeftCell, 2)
    Next areaItem

End Sub

Sub createControlTable()
    'move to sheet that corresponds to this control sheet
    Dim inUpLeftCell As Range
    Dim rowStart As String, rowEnd As String
    Dim colStart As String, colEnd As String
    Dim sampleRow As Integer
    Dim clw As New CellWorker
    Dim tmpStr As String, tmpArr As Variant
    Dim tmpRng As Range
    Dim rowVal As String, colVal As String
    
    sampleRow = 1
    Sheets("control_table_ÁÀÐ_ø").Select
    Set inUpLeftCell = Range("A3")
    Do While inUpLeftCell.value <> 0
        'Debug.Assert inUpLeftCell.value <> ""
        If InStr(1, inUpLeftCell.value, ":") <> 0 Or InStr(1, Cells(1, inUpLeftCell.Column).value, ":") <> 0 Then
            tmpArr = Split(inUpLeftCell.value, ":")
            tmpStr = tmpArr(0)
            If UBound(tmpArr) > 0 Then
                rowStart = computeRowOrCol(tmpArr(0), True)
                rowEnd = computeRowOrCol(tmpArr(1), True)
            Else
                rowStart = computeRowOrCol(tmpStr, True)
                rowEnd = computeRowOrCol(tmpStr, True)
            End If
            Debug.Assert InStr(1, rowStart, "R") <> 0 And InStr(1, rowStart, "C") = 0
            Set tmpRng = clw.move_right(inUpLeftCell, 2)
            Do While Cells(sampleRow, tmpRng.Column).value <> ""
                tmpArr = Split(Cells(sampleRow, tmpRng.Column).value, ":")
                tmpStr = tmpArr(0)
                If UBound(tmpArr) > 0 Then
                    colStart = computeRowOrCol(tmpArr(0))
                    colEnd = computeRowOrCol(tmpArr(1))
                Else
                    colStart = computeRowOrCol(tmpStr)
                    colEnd = computeRowOrCol(tmpStr)
                End If
                Debug.Assert InStr(1, colStart, "C") <> 0
                tmpStr = rowStart & colStart & ":" & rowEnd & colEnd
                tmpRng.value = tmpStr
                Set tmpRng = clw.move_right(tmpRng, 2)
            Loop
        Else
            Do While Cells(sampleRow, tmpRng.Column).value <> ""
        
                rowStart = computeRowOrCol(tmpStr, True)
                rowEnd = computeRowOrCol(tmpStr, True)
                colStart = computeRowOrCol(tmpStr)
                colEnd = computeRowOrCol(tmpStr)
                tmpStr = rowStart & colStart & ":" & rowEnd & colEnd
                tmpRng.value = tmpStr
                Set tmpRng = clw.move_right(tmpRng, 2)
            Loop
        
        End If
        Set inUpLeftCell = clw.move_down(inUpLeftCell, 2)
    Loop
    
    
End Sub
Private Function computeRowOrCol(addr As Variant, Optional rowAddr As Boolean) As String
    'Takes addr in R1C1 notation and if rowAddr is true returns R[] part otherwise C[] part
    '>>>computeRowOrCol("R[10]C[5]",True)
    '"R[10]"
    '>>>computeRowOrCol("R[10]C[5]")
    '"C[5]"
    
    Dim i As Integer
    Dim tmpStr As String

    i = InStr(1, addr, "]")
    Debug.Assert i <> 0
    
    If rowAddr Then
        tmpStr = Left(addr, i)
        computeRowOrCol = tmpStr
    Else
        tmpStr = Right(addr, Len(addr) - i)
        computeRowOrCol = tmpStr
    End If
    
End Function


