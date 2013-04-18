Attribute VB_Name = "Controller"
'module that will control interaction between view(copyMineUF) and model(bbUgol_copyPaste)
Dim srcWB As Workbook, destWB As Workbook
Dim ctrlGenSht As Worksheet, cmbxCondSht As Worksheet
Dim workRangeUpLeftCell As Range
Dim constValColl As Collection
Public techChange As Boolean 'flag that used for turning off combobox's change events
Dim calcType As Variant

Private Sub initialize_constValColl()
'''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'entire application's constant value initialization



    'don't change values of this collection in your code. Should be replaced by some immutable type, but i don't
    'know what I should use in excel
    Set constValColl = New Collection
    constValColl.Add "M1", "tmp_filter_output" 'temporary output from filter. This output will be used as workrange
    'for copyProc
    constValColl.Add "O2", "workRangeUpLeftCell" 'first cell that should contain values of relToRange
    'from tmp_filter_output
    constValColl.Add "control_table_general", "ctrlGenShtName"
    constValColl.Add "cmbx_condition_sht", "cmbxCondShtName"
    constValColl.Add ActiveWorkbook.Name, "destWBName"
    constValColl.Add "control_table_", "sht_control_table_prefix"
    constValColl.Add "A1", "upLeftCell_for_ctrl_sht"
    constValColl.Add "mineMan", "mine_management_prefix"
    constValColl.Add "mine", "mine_prefix"
    constValColl.Add "CmBx", "cmbx_postfix"
    constValColl.Add "Lbl", "lbl_postfix"
    
End Sub
Sub startNewCopyMine_click()
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
    'entry point for application
    Call initialize_constValColl
    Debug.Assert Not constValColl Is Nothing And constValColl.Count > 0
    Set destWB = Workbooks(constValColl("destWBName"))
    Set ctrlGenSht = destWB.Sheets(constValColl("ctrlGenShtName"))
    Set cmbxCondSht = destWB.Sheets(constValColl("cmbxCondShtName"))
    Call unhide_everything
    copyMineUF.Show
End Sub
'Dim workRange As Range
Sub copyBtnClicked()
    Dim shtName As String
    Dim tmpRng As Range
    Dim clw As New CellWorker
    
    ctrlGenSht.Activate
    If workRangeUpLeftCell Is Nothing Then
        Set workRangeUpLeftCell = Range(constValColl("workRangeUpLeftCell"))
    End If
    Set tmpRng = workRangeUpLeftCell
    Do While tmpRng.value <> ""
        ctrlGenSht.Activate
        shtName = Cells(1, tmpRng.Column)
        Debug.Assert shtName <> ""
        Call bbUgol_copyPaste.copyProc(shtName, tmpRng.value, constValColl)
        Set tmpRng = clw.move_down(tmpRng)
    Loop
        
    Set tmpRng = Nothing
    Set tmpRng = clw.move_right(workRangeUpLeftCell)
    'loop that helps to skip some sheets
    Do While tmpRng.value = "" And Cells(1, tmpRng.Column).value <> ""
        Set tmpRng = clw.move_right(tmpRng)
    Loop
    If tmpRng.value <> "" Then
        Set workRangeUpLeftCell = tmpRng
        Call copyBtnClicked 'recursive call
    End If
    
End Sub


Function proccesFileSelection() As String

    'test
    Dim Filt As String
    Dim FilterIndex As Integer
    'Dim FileName As Variant
    Dim Title As String
    Dim flw As New FileWorker
    Dim tmpStr As String
    Dim cachedWb As Workbook
    '@todo create file opener
    
    ' Set up list of file filters
    Filt = "Excel files (*.xlsx;*.xltx;*.xlsm;*.xltm),*.xlsx;*.xltx;*.xlsm;*.xltm"
    FilterIndex = 1
    ' Set the dialog box caption
    Title = "Выберите файлы для консолидации"
    ' Get the file name
    
    FileName = Application.GetOpenFilename _
                (FileFilter:=Filt, _
                FilterIndex:=FilterIndex, _
                Title:=Title, _
                MultiSelect:=False)
    ' Exit if dialog box canceled
    If FileName = False Then
        MsgBox "Пожалуйста, выберите файл"
        Exit Function
    End If
    tmpStr = CStr(FileName)
    proccesFileSelection = tmpStr
    'this should be moved to userform code
    copyMineUF.srcNameLbl.ForeColor = vbBlack
    copyMineUF.srcNameLbl.ControlTipText = "Имя файла из которого будет выполнятся копирование"
    copyMineUF.mineManLbl.ForeColor = vbRed
    Set cachedWb = ActiveWorkbook
    
    Workbooks.Open tmpStr, False, True
    
    constValColl.Add flw.extractNameWithExt(tmpStr), "srcWBName" 'this should be reassigned accrodingly to name of opened file
    
    cachedWb.Activate

End Function


Sub unloadCopyMineUF()
    
    
    Unload copyMineUF
    Call cmbx_cleaning
    Call cmbx_cond_cleaning
    Call tmpFilterRegionClear
    Call hide_everything
    
    Set destWB = Nothing
    Set ctrlGenSht = Nothing
    Set cmbxCondSht = Nothing
    Set workRangeUpLeftCell = Nothing
    Set constValColl = Nothing
    'Application.Calculate
    
End Sub



Sub generalFiltering()
    '@todo replace by something more dynamic
    cmbxCondSht.Range("A2").value = copyMineUF.mineManCmBx.Text
    cmbxCondSht.Range("B2").value = copyMineUF.mineCmBx.Text
    Call tmpFilterRegionClear
    Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=cmbxCondSht.Range("A1:B2"), CopyToRange:=Range(constValColl("tmp_filter_output"))
    'Set generalFiltering = Range("M1").CurrentRegion
    'ctrlGenSht.Activate

End Sub


Private Sub cmbx_cleaning()
    techChange = True
    With copyMineUF
        .mineCmBx.Text = ""
        .mineManCmBx.Text = ""
    End With
    techChange = False
End Sub


Private Sub cmbx_cond_cleaning()

    cmbxCondSht.Activate
    Range("F1").CurrentRegion.Clear
    Range("H1").CurrentRegion.Clear
    Range("J1").CurrentRegion.Clear
    Range("A2:B2").Clear
    
End Sub

Private Sub tmpFilterRegionClear()
    'M1 is cell address where tmp filter stores its output
    techChange = True
    ctrlGenSht.Activate
    Range(constValColl.Item("tmp_filter_output")).CurrentRegion.Clear
    techChange = False
End Sub
Function computerRowSource(cmbxName As String) As String

    Dim tmpStr As String
    
    If cmbxName = "mineManCmBx" Then
        tmpStr = mineManCmBx_compute_rowsource
    End If
    If cmbxName = "mineCmBx" Then
        tmpStr = mineCmBx_compute_rowsource
    End If

    computerRowSource = tmpStr
End Function
Private Function mineManCmBx_compute_rowsource() As String
'
    Dim tmpStr As String
'
    ctrlGenSht.Select
    Range(Range("A1"), Range("A1").End(xlDown)).Copy
    cmbxCondSht.Select
    Range("F1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        Range("D1:D2"), CopyToRange:=Range("H1"), Unique:=True
    Range(Range("H2"), Range("H2").End(xlDown)).Select
    tmpStr = Selection.Parent.Name & "!" & Selection.Address(False, False)
    mineManCmBx_compute_rowsource = tmpStr
End Function



Private Function mineCmBx_compute_rowsource() As String
    Dim tmpStr As String
    
    ctrlGenSht.Activate
    Range(Range("N2"), Range("N2").End(xlDown)).Select
    Selection.Copy
    cmbxCondSht.Activate
    Range("J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    tmpStr = Selection.Parent.Name & "!" & Selection.Address(False, False)

    mineCmBx_compute_rowsource = tmpStr

End Function

Private Sub hide_everything()

    ctrlGenSht.Visible = xlSheetVeryHidden
    cmbxCondSht.Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = calcType
    
End Sub

Private Sub unhide_everything()
    calcType = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ctrlGenSht.Visible = xlSheetVisible
    cmbxCondSht.Visible = xlSheetVisible
    
End Sub
