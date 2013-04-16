Attribute VB_Name = "bbUgol_copyPaste"
Dim relToRange  As Range
Dim addrColl As Collection
Dim srcWB As Workbook, destWB As Workbook


Sub copyProc(shtName As String, relToRngAddr As String, constValColl As Collection)

    'constValColl should contain at least source and destination workbook name
    'prefix that is being added to shtName for ctrlShtName creation
    'upLeftCell for ctrlSht that contains first address for copy
    Dim ctrlSht As Worksheet
    Dim ctrlRng As Range
    Dim addrForCopy As String
    
    Set srcWB = Workbooks(constValColl("srcWBName"))
    Set destWB = Workbooks(constValColl("destWBName"))
    
    Set srcWSht = srcWB.Sheets(shtName)
    Set ctrlSht = destWB.Sheets(constValColl("sht_control_table_prefix") & shtName)
    Set destWSht = destWB.Sheets(shtName)
    
    Set ctrlRng = ctrlSht.Range(constValColl("upLeftCell_for_ctrl_sht")) 'upLeftCell for mine range
    Set relToRange = ctrlSht.Range(relToRngAddr)
    
    'important
    ctrlSht.Visible = xlSheetVisible
    ctrlSht.Activate
    
    Call moveThroughRows(ctrlRng)
    
    ctrlSht.Visible = xlSheetVeryHidden
    
    'Copy one range to another
    For Each addr In addrColl
        Call copyRange(addr)
    Next addr
    
End Sub

Sub unhide_everything(disableAppOperations As Boolean)
    If disableAppOperations Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    End If
    ctrlSht.Visible = xlSheetVeryHidden

End Sub

Sub hide_everything()
    ctrlSht.Visible = xlSheetVeryHidden
End Sub

'Open files for copy

'Find needed mine or it's range


'Compute range address for copying
Private Sub moveThroughRows(inRange As Range)
    'Procedure moves through all most left non-empty cells in rows
    Dim nextRowRange As Range
    Dim clw As New CellWorker
     
    
    Call processRowOfRanges(inRange)
    
    Set nextRowRange = clw.move_down(inRange, 2)
    
    If nextRowRange.value <> "" Then
        Call moveThroughRows(nextRowRange)
    End If

End Sub

Private Sub processRowOfRanges(inRange As Range)

    Dim addrForProc As String
    Dim clw As New CellWorker
    Dim nextRange As Range
    
    'range address converted to A1 notation
    addrForProc = convertToA1(inRange.value)
    
    If addrColl Is Nothing Then
        Set addrColl = New Collection
    End If
    
    addrColl.Add (addrForProc)
    
    'moves to next range
    Set nextRange = clw.move_right(inRange, 2)
        
    If nextRange.value <> "" Then
        Call processRowOfRanges(nextRange)  ' - recursive call
    End If
End Sub

Private Sub copyRange(addrForCopy As Variant)

    destWSht.Range(addrForCopy).Value2 = srcWSht.Range(addrForCopy).Value2

End Sub


Private Function convertToA1(inRange As String) As String
    '
    '(str)->str
    
    'Returns converted inRange address to xlA1 style
    
    'relToRange="E149"
    '>>>convertToA1(R[6]C[2]:R[7]C[3])
    '"G155:H156"
    convertToA1 = Application.ConvertFormula(inRange, xlR1C1, xlA1, , relToRange)
    
End Function



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


''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''Helpers'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''


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
    
    Set relToRange = Range("E149")
    Set upLeftCell = Sheets("control_table_" & ActiveSheet.Name).Range("I44")
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
    
    Set relToRange = Range("E149")
    Set upLeftCell = Sheets("control_table_" & ActiveSheet.Name).Range("I10")
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
    
    sampleRow = 10
    Sheets("control_table_ÁÏÑÑ_ø").Select
    Set inUpLeftCell = Range("I12")
    Do While inUpLeftCell.value <> 0
        Debug.Assert inUpLeftCell.value <> "R[74]C[2]:R[75]C[2]"
        If InStr(1, inUpLeftCell.value, ":") <> 0 Then
            tmpArr = Split(inUpLeftCell.value, ":")
            tmpStr = tmpArr(0)
            rowStart = computeRowOrCol(tmpArr(0), True)
            rowEnd = computeRowOrCol(tmpArr(1), True)
            Debug.Assert InStr(1, rowStart, "R") <> 0 And InStr(1, rowStart, "C") = 0
            Set tmpRng = clw.move_right(inUpLeftCell, 2)
            Do While Cells(sampleRow, tmpRng.Column).value <> ""
                tmpArr = Split(Cells(sampleRow, tmpRng.Column).value, ":")
                tmpStr = tmpArr(0)
                colStart = computeRowOrCol(tmpArr(0))
                colEnd = computeRowOrCol(tmpArr(1))
                Debug.Assert InStr(1, colStart, "C") <> 0
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

Sub checkMineRange()
    Sheets("control_table_ÁÏÑÑ_ø").Select
    'Set relToRange = Range("E287")
    Application.EnableEvents = False
    Call processMineRange(Range("A1"))
    Application.EnableEvents = True
End Sub
