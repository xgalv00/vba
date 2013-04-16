Attribute VB_Name = "Controller"
'module that will control interaction between view(copyMineUF) and model(bbUgol_copyPaste)
Dim srcWB As Workbook, destWB As Workbook
Dim ctrlGenSht As Worksheet, cmbxCondSht As Worksheet
Dim workRangeUpLeftCell As Range
Dim constValColl As Collection
Public techChange As Boolean 'flag that used for turning off combobox's change events

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
    constValColl.Add "model_in.xlsm", "srcWBName" 'this should be reassigned accrodingly to name of opened file
    constValColl.Add "control_table_", "sht_control_table_prefix"
    constValColl.Add "A1", "upLeftCell_for_ctrl_sht"
    
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
    copyMineUF.Show False
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
        shtName = Cells(1, tmpRng.Column)
        Call bbUgol_copyPaste.copyProc(shtName, tmpRng.value, constValColl)
        Set tmpRng = clw.move_down(tmpRng)
    Loop
        
    Set tmpRng = Nothing
    Set tmpRng = clw.move_right(workRangeUpLeftCell)
    If tmpRng.value <> "" Then
        Set workRangeUpLeftCell = tmpRng
        Call copyBtnClicked 'recursive call
    End If
    
End Sub

Sub copyStyleChkBxClicked()
    '@todo
End Sub

Sub mineCmBx_Changed()

    Call enable_copyBtn
    copyMineUF.mineLbl.ForeColor = vbBlack
    Call generalFiltering

End Sub
Sub mineManCmBx_Changed()
    Call generalFiltering
    
    copyMineUF.mineManLbl.ForeColor = vbBlack

    If copyMineUF.copyStyleChkBx Then
        copyMineUF.mineCmBx.Enabled = True
        copyMineUF.mineLbl.ForeColor = vbRed
        copyMineUF.mineCmBx.RowSource = Controller.computerRowSource("mineCmBx")
    Else
        Call enable_copyBtn
    End If

End Sub

Sub proccesFileSelection()

    'test
    Dim Filt As String
    Dim FilterIndex As Integer
    'Dim FileName As Variant
    Dim Title As String
    Dim flw As New FileWorker
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
        Exit Sub
    End If
    
    copyMineUF.srcNameLbl.Caption = CStr(FileName)
    
    copyMineUF.srcNameLbl.ForeColor = vbBlack
    copyMineUF.srcNameLbl.ControlTipText = "Имя файла из которого будет выполнятся копирование"
    copyMineUF.mineManLbl.ForeColor = vbRed
    
    copyMineUF.mineManCmBx.Enabled = True
    copyMineUF.mineManCmBx.RowSource = computerRowSource("mineManCmBx")


End Sub


Sub unloadCopyMineUF()
    
    
    Unload copyMineUF
    Call cmbx_cleaning
    Call cmbx_cond_cleaning
    Call tmpFilterRegionClear
    
    Set destWB = Nothing
    Set ctrlGenSht = Nothing
    Set cmbxCondSht = Nothing
    Set workRangeUpLeftCell = Nothing
    
End Sub

Private Sub enable_copyBtn()
    copyMineUF.copyBtn.Enabled = True
End Sub

Private Sub generalFiltering()
    
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
Private Function computerRowSource(cmbxName As String) As String

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

