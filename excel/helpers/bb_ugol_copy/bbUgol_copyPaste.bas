Attribute VB_Name = "bbUgol_copyPaste"
Dim relToRange  As Range
Dim addrColl As Collection
Dim srcWB As Workbook, destWB As Workbook
Dim ctrlSht As Worksheet


Sub copyProc(shtName As String, relToRngAddr As String, constValColl As Collection)

    'constValColl should contain at least source and destination workbook name
    'prefix that is being added to shtName for ctrlShtName creation
    'upLeftCell for ctrlSht that contains first address for copy
    Dim ctrlRng As Range
    Dim addrForCopy As String
    
    Set addrColl = Nothing
    Set srcWB = Workbooks(constValColl("srcWBName"))
    Set destWB = Workbooks(constValColl("destWBName"))
    
    'Set srcWSht = srcWB.Sheets(shtName)
    Set ctrlSht = destWB.Sheets(constValColl("sht_control_table_prefix") & shtName)
    'Set destWSht = destWB.Sheets(shtName)
    
    Set ctrlRng = ctrlSht.Range(constValColl("upLeftCell_for_ctrl_sht")) 'upLeftCell for mine range
    Set relToRange = ctrlSht.Range(relToRngAddr)
    
    'important
    Call unhide_everything
    ctrlSht.Activate
    
    Call moveThroughRows(ctrlRng)
    
    Call hide_everything
    
    'Copy one range to another
    For Each addr In addrColl
        Call copyRange(shtName, addr)
    Next addr
    
End Sub

Private Sub unhide_everything(Optional disableAppOperations As Boolean)
    If disableAppOperations Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    End If
    ctrlSht.Visible = xlSheetVisible

End Sub

Private Sub hide_everything()
    ctrlSht.Visible = xlSheetVeryHidden
End Sub

'Open files for copy


'Compute range address for copying
Sub moveThroughRows(inRange As Range)
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

Private Sub copyRange(shtName As String, addrForCopy As Variant)

    destWB.Sheets(shtName).Range(addrForCopy).Value2 = srcWB.Sheets(shtName).Range(addrForCopy).Value2

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





Sub checkMineRange()
    Dim tmpSht As Worksheet
    Set addrColl = Nothing
    Set tmpSht = Sheets("control_table_Б_пр_во")
    tmpSht.Visible = xlSheetVisible
    tmpSht.Select
    Set relToRange = Range("E128")
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
