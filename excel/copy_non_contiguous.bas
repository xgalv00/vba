Attribute VB_Name = "copy_non_contiguous"
Sub an_copy_test()

    Dim srcWBook As New Workbook, dstWBook As New Workbook
    Dim srcSht As New Worksheet, dstSht As New Worksheet
    Dim numRows As Integer, numCols As Integer
    Dim upLeftCell As Range, downRigCell As Range
    
    Set dstWBook = Workbooks("test.xlsx")
    Set srcWBook = Workbooks("RN_BILLS1.xlsm")
    Set srcSht = srcWBook.Sheets(1)
    Set dstSht = dstWBook.Sheets(1)
    
    srcSht.Activate
    srcSht.UsedRange.Select
    numRows = srcSht.UsedRange.Rows.Count
    numCols = srcSht.UsedRange.Columns.Count
    Selection.SpecialCells(xlCellTypeVisible).Select
    Call CopyMultipleSelection
    
    
End Sub

Sub CopyMultipleSelection()
    Dim SelAreas() As Range
    Dim PasteRange As Range
    Dim UpperLeft As Range
    Dim NumAreas As Long, i As Long
    Dim TopRow As Long, LeftCol As Long
    Dim RowOffset As Long, ColOffset As Long
    Dim tmpRow As Long, tmpCol As Long
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    NumAreas = Selection.Areas.Count
    ReDim SelAreas(1 To NumAreas)
    For i = 1 To NumAreas
        Set SelAreas(i) = Selection.Areas(i)
    Next
    
    TopRow = ActiveSheet.Rows.Count
    LeftCol = ActiveSheet.Columns.Count
    For i = 1 To NumAreas
        If SelAreas(i).Row < TopRow Then TopRow = SelAreas(i).Row
        If SelAreas(i).Column < LeftCol Then LeftCol = SelAreas(i).Column
    Next
    Set UpperLeft = Cells(TopRow, LeftCol)
    
    On Error Resume Next
    Workbooks("test.xlsx").Activate
    Set PasteRange = Range("A1")
    PasteRange.Select
    'Set PasteRange = Application.InputBox _
    '(Prompt:="Specify the upper-left cell for the paste range:", _
    'Title:="Copy Multiple Selection", _
    'Type:=8)
    On Error GoTo 0
    
    If TypeName(PasteRange) <> "Range" Then Exit Sub
    
    'Set PasteRange = PasteRange.Range("A1")
    Workbooks("RN_BILLS1.xlsm").Activate
    
    For i = 1 To NumAreas
        'RowOffset = SelAreas(i).Row - TopRow
        'ColOffset = SelAreas(i).Column - LeftCol
        Workbooks("RN_BILLS1.xlsm").Activate
        SelAreas(i).Select

        
        'Cells(tmpRow, tmpCol).Select
        If tmpRow < Selection.Row + SelAreas(i).Rows.Count - 1 Then
            RowOffset = Workbooks("test.xlsx").Sheets(1).UsedRange.Rows.Count
            ColOffset = 0 'replace by variable
        Else
            ColOffset = Workbooks("test.xlsx").Sheets(1).UsedRange.Columns.Count
            RowOffset = tmpRow - Selection.Row - SelAreas(i).Rows.Count + 1
        End If
        'Debug.Print "Area num " & i & " end row " & SelAreas(i).Rows.Count & " end column " & SelAreas(i).Columns.Count
        SelAreas(i).Select
        'RowOffset = 14
        'ColOffset = 0
        
        'last row and column
        tmpRow = Selection.Row + SelAreas(i).Rows.Count - 1
        tmpCol = SelAreas(i).Column + SelAreas(i).Columns.Count - 1
        'Debug.Print Selection.EntireRow.Hidden
        SelAreas(i).Copy
        PasteRange.Offset(RowOffset, ColOffset).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        Application.CutCopyMode = False
    Next i
End Sub
