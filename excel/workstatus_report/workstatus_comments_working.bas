Attribute VB_Name = "workstatus_comments_working"
Public technicalChange As Boolean
Dim statusWB As Workbook
Dim wStatSht As Worksheet, comDraftSht As Worksheet, wStatDraftSht As Worksheet
Dim msfoTableSht As Worksheet, usrTableSht As Worksheet
Dim workRange As Range

Sub changeStatus()

    bulkStatusChangeUF.Show

End Sub


Sub refreshSht()

    Application.Run "MNU_eSUBMIT_REFSCHEDULE_SHEET_REFRESH"

End Sub


Sub refresh()

    Application.Run "MNU_eTOOLS_REFRESH"

End Sub

Sub readComments()
    
    Dim tmpCell As Range
    Dim clw As New CellWorker
    Dim endRow As Integer, endCol As Integer
    Dim i As Integer, j As Integer
    
    Call initialize_WS_variables
    
    'Application.ScreenUpdating = False
    Application.EnableEvents = False
    comDraftSht.Activate
    
    endRow = workRange.Row + workRange.Rows.Count
    endCol = workRange.Column + workRange.Columns.Count
       
    For i = Range("N11").Row To endRow - 1
    
        For j = Range("N11").Column To endCol - 1
            Set tmpCell = Cells(i, j)
            tmpCell.Select 'test line
            If tmpCell.Value <> "" Then
                'Debug.Print tmpCell.Value
                
                wStatSht.Activate
                If Range(tmpCell.Address).Comment Is Nothing Then
                    Range(tmpCell.Address).AddComment
                End If
                Range(tmpCell.Address).Comment.text text:=tmpCell.Value
                comDraftSht.Activate
                
            End If
        Next j
    
    Next i
    
    wStatSht.Activate
    Application.EnableEvents = True
    'Application.ScreenUpdating = True
End Sub

Sub writeComments()

    Dim tmpCell As Range
    Dim clw As New CellWorker
    Dim endRow As Integer, endCol As Integer
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    
    Call initialize_WS_variables
    
    'Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    wStatSht.Activate
    
    endRow = workRange.Row + workRange.Rows.Count
    endCol = workRange.Column + workRange.Columns.Count
       
    For i = Range("N11").Row To endRow - 1
    
        For j = Range("N11").Column To endCol - 1
            Set tmpCell = Cells(i, j)
            tmpCell.Select 'test line
            If Not tmpCell.Comment Is Nothing Then
                'Debug.Print tmpCell.Value
                
                tmpStr = Range(tmpCell.Address).Comment.text
                comDraftSht.Activate
                Cells(tmpCell.Row, (tmpCell.Column + workRange.Columns.Count)).Value = tmpStr
                Cells(tmpCell.Row, (tmpCell.Column + workRange.Columns.Count)).Select
                wStatSht.Activate
                tmpStr = ""
                
            End If
        Next j
    
    Next i
    Application.EnableEvents = True
   'Application.ScreenUpdating = True
End Sub

Sub create_comboBxs()

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Helper!$B$1:$B$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub tweak_macro()
'
' tweak_macro1 Макрос
'

'
    technicalChange = True
    
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "=EVLCK(R1C2,0,R3C13,RC8,R9C)"
    '@todo rewrite  this autofill with use of workrange
    Selection.AutoFill Destination:=Range("N11:N146"), Type:=xlFillDefault
    Range("N11:N146").Select
    Selection.AutoFill Destination:=Range("N11:AY146"), Type:=xlFillDefault
    Range("N11:AY146").Select
    Range("N11").Select
    
    technicalChange = False
End Sub
Sub clear_statuses(Optional shtForClear As Worksheet)
    Call initialize_WS_variables
    
    technicalChange = True
    If shtForClear Is Nothing Then
        Range(workRange.Address(False, False)).ClearContents
    Else
        shtForClear.Range(workRange.Address(False, False)).ClearContents
    End If
    Range(workRange.Address(False, False)).ClearComments
    technicalChange = True
    
End Sub


Sub initialize_WS_variables()
Attribute initialize_WS_variables.VB_ProcData.VB_Invoke_Func = " \n14"
'
' test Макрос
'
'    Dim clw As New CellWorker
'
    Set statusWB = ActiveWorkbook
    
    Set wStatSht = statusWB.Sheets("WorkStatus")
    Set wStatDraftSht = statusWB.Sheets("WorkStatusDraft")
    Set comDraftSht = statusWB.Sheets("CommentsDraft")
    Set usrTableSht = statusWB.Sheets("user_table")
    Set msfoTableSht = statusWB.Sheets("msfo_table")
    'range address of statuses
    Set workRange = calcWorkRange
    'workRange.Select
    
End Sub

Function calcWorkRange() As Range

    Dim colKeyRange As String
    Dim rowKeyRange As String
    Dim tmpArr As Variant
    Dim tmpString As String
    Dim tmpRange As Range
    Dim upLeftAddr As String
    Dim downRightAddr As String
    
    wStatSht.Activate
    colKeyRange = Range("B34").Value
    tmpArr = Split(colKeyRange, "$")
    upLeftAddr = tmpArr(1) & "11"
    tmpString = tmpArr(3)
    rowKeyRange = Range("B35").Value
    tmpArr = Split(rowKeyRange, "$")
    downRightAddr = tmpString & tmpArr(4)
    
    Set calcWorkRange = Range(upLeftAddr & ":" & downRightAddr)
End Function

Sub prepareWorkspace()
    Dim colorColl As New Collection
    Dim keyColl As New Collection
    Call initialize_WS_variables
    Call unhide_everything
    
    wStatSht.Activate
    Range("N11").Select
    Call copyStatuses
    
    workRange.Select
    Call create_comboBxs
    Selection.Locked = False
    
    Call readComments
    Call hide_everything


End Sub

Private Function Pass(sh)
'
' Прочитать пароль на листе
'
    ' Поиск ячейки-маркера
    Set f = sh.Cells.find("PasswordBPC", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        Set f = sh.Cells(f.Row + 1, f.Column)
        Pass = f.Value
        If sh.ProtectContents = False Then
            f.NumberFormat = ";;;"
            'f.Locked = True
            f.FormulaHidden = True
            Set r = Range(sh.Cells(f.Row - 1, f.Column), f)
            r.Interior.ThemeColor = xlThemeColorAccent1
            r.Interior.TintAndShade = 0.4
        End If
    End If
End Function


Sub unhide_everything()
    Call initialize_WS_variables
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wStatSht.Unprotect Pass(wStatSht)
    comDraftSht.Visible = xlSheetVisible
    Sheets("Helper").Visible = xlSheetVisible
    wStatDraftSht.Visible = xlSheetVisible
    usrTableSht.Visible = xlSheetVisible
    msfoTableSht.Visible = xlSheetVisible

End Sub

Sub hide_everything()
    Call initialize_WS_variables
    comDraftSht.Visible = xlSheetVeryHidden
    Sheets("Helper").Visible = xlSheetVeryHidden
    wStatDraftSht.Visible = xlSheetVeryHidden
    usrTableSht.Visible = xlSheetVeryHidden
    msfoTableSht.Visible = xlSheetVeryHidden
    wStatSht.Protect Pass(wStatSht)
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub sendComments()
    Call initialize_WS_variables
    Call unhide_everything
    
    Call writeComments
    
    Call hide_everything
    'wStatSht.Select
    'workRange.Select
    'Call clearComboBxs
    Application.Run "MNU_eSUBMIT_REFSCHEDULE_BOOK_NOACTION_SHOWRESULT"
    'workRange.Select
    'Call create_comboBxs
End Sub

Sub clearComboBxs()
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub copyStatuses()
    
    'copy status' filling from draft to clean copy
    Dim clw As New CellWorker
    Dim eng_rus_dict As Collection
    Dim firstCellInRow As Range
    Dim tmpRng As Range
    Dim convVal As String
    
    technicalChange = True
    
    Set eng_rus_dict = ws_change_module.make_dictionary()
    wStatDraftSht.Select
    Set firstCellInRow = Range("N11")
    
    Do While firstCellInRow.Value <> ""
        Set tmpRng = firstCellInRow
        Do While tmpRng.Value <> "" And tmpRng.Value <> "#ERR"
            'translates values
            wStatSht.Range(tmpRng.Address(False, False)).Value = eng_rus_dict(wStatDraftSht.Range(tmpRng.Address(False, False)).Value)
            Set tmpRng = clw.move_right(tmpRng)
        Loop
        Set firstCellInRow = clw.move_down(firstCellInRow)
    Loop
    'Selection.Copy
    'workRange.Copy
    wStatSht.Select

    
    technicalChange = False

End Sub

Function isInWorkrange(cellToExam As Range) As Boolean

    Dim workRangeAddr As String
    Dim upLeftCell As Range, downRightCell As Range
    Dim tmpArray As Variant
    
    isInWorkrange = False
    wStatSht.Activate
    workRangeAddr = workRange.Address(False, False)
    tmpArray = Split(workRangeAddr, ":")
    Set upLeftCell = Range(tmpArray(0))
    Set downRightCell = Range(tmpArray(1))
    
    If cellToExam.Row <= downRightCell.Row And cellToExam.Column <= downRightCell.Column And cellToExam.Row >= upLeftCell.Row And cellToExam.Column >= upLeftCell.Column Then
        isInWorkrange = True
    End If

End Function

Sub bulkStatusChange(statusVal)

    'Procedure for bulk status change
    
    Dim cellsCol As Collection
    Dim rangeAddrToExam As String
    Dim cellToExam As Range
    Dim upLeftCell As Range
    Dim clw As New CellWorker
    
    Call initialize_WS_variables
    wStatSht.Activate
    
    Set cellsCol = New Collection
    
    If Selection.Areas.Count = 1 Then
        rangeAddrToExam = Selection.Areas(1).Address(False, False)
        Set cellToExam = Cells(Selection.Areas(1).Row, Selection.Areas(1).Column)
        Set upLeftCell = cellToExam
        Do While isInRange(rangeAddrToExam, cellToExam)
            Do While isInRange(rangeAddrToExam, cellToExam)
                cellsCol.Add cellToExam.Address(False, False)
                Set cellToExam = clw.move_down(cellToExam)
            Loop
            Set cellToExam = clw.move_right(upLeftCell)
        Loop
    ElseIf Selection.Areas.Count > 1 Then
    
    End If
    'Create collection with cell addresses to process
    
    'call record_change for each cell in collection
    
    'collect message and send it

End Sub

Function isInRange(rangeAddrToExam As String, cellToExam As Range) As Boolean

    Dim workRangeAddr As String
    Dim upLeftCell As Range, downRightCell As Range
    Dim tmpArray As Variant
    
    isInRange = False
    workRangeAddr = rangeAddrToExam
    tmpArray = Split(workRangeAddr, ":")
    Set upLeftCell = Range(tmpArray(0))
    Set downRightCell = Range(tmpArray(1))
    
    If cellToExam.Row <= downRightCell.Row And cellToExam.Column <= downRightCell.Column And cellToExam.Row >= upLeftCell.Row And cellToExam.Column >= upLeftCell.Column Then
        isInRange = True
    End If

End Function

Function record_change(changeCellAddr As String) As Boolean
    Dim srcWSht As Worksheet ', destWSht As Worksheet
    Dim srcCellFormula As String ', destCellFormula As String
    Dim compVal As String, dsVal As String, timeVal As String, statusVal As String
    Dim compValAddr As String
    Dim tmpArray As Variant
    
    'Call unhide_everything
    Set srcWSht = Sheets("WorkStatusDraft")
    srcWSht.Select
    'Set destWSht = Sheets("Changed1")
    srcCellFormula = srcWSht.Range(changeCellAddr).Formula
    'destWSht.Activate
    'Range(changeCellAddr).Formula = srcCellFormula
    
    tmpArray = Split(srcCellFormula, ",")
    compValAddr = Left(tmpArray(4), Len(tmpArray(4)) - 1)
    timeVal = Range(tmpArray(2)).Value
    dsVal = Range(tmpArray(3)).Value
    compVal = Range(compValAddr).Value
    Sheets("WorkStatus").Activate
    statusVal = Range(changeCellAddr).Value
    Debug.Assert compVal <> "" Or dsVal <> "" Or timeVal <> "" Or statusVal <> ""
    Call wsChangePrep(compVal, dsVal, timeVal, statusVal)
    'Call hide_everything
    If ws_change_module.statusChanged Then
        record_change = True
    End If
End Function

Function isAuthorized(changedCell As Range) As String
    Dim compName As String
    Dim periodVal As String
    
    compName = Sheets("WorkStatus").Cells(10, changedCell.Column).Value
    periodVal = Sheets("WorkStatus").Range("N3")
    
'    If Not isCompanyInUsrCompColl(compName) Then
'        Exit Function
'    End If
'    If Not isUsrHasApprType(changedCell.Value) Then
'        Exit Function
'    End If
    If Not isCompanyInUsrCompColl(compName) Then
        isAuthorized = "у Вас недостаточно прав для изменения статуса по компании " & compName
        Exit Function
    ElseIf Not isUsrHasApprType(changedCell.Value) Then
        isAuthorized = "у Вас недостаточно прав, чтобы поставить статус " & changedCell.Value
        Exit Function
    ElseIf Not isValidPeriod(periodVal) Then
        isAuthorized = "необходимо выбрать месяц в поле период"
        Exit Function
    End If
    
    isAuthorized = "ok"
End Function
Function isValidPeriod(periodVal As String) As Boolean

    Dim monthCol As New Collection
    Dim monthArg As String
    Dim tmpArray As Variant
    
    'this used because of language issues (month naming)
    monthCol.Add "Январь"
    monthCol.Add "Февраль"
    monthCol.Add "Март"
    monthCol.Add "Апрель"
    monthCol.Add "Май"
    monthCol.Add "Июнь"
    monthCol.Add "Июль"
    monthCol.Add "Август"
    monthCol.Add "Сентябрь"
    monthCol.Add "Октябрь"
    monthCol.Add "Ноябрь"
    monthCol.Add "Декабрь"

    tmpArray = Split(periodVal, " ")
    monthArg = tmpArray(0)
    
    isValidPeriod = isExistInCol(monthArg, monthCol)
End Function

Private Function isExistInCol(itemVal As String, colForSearch As Collection) As Boolean
    'loop through given collection and if meet given value return true
    Dim resBool As Boolean
    resBool = False
    itemVal = LCase(Trim(itemVal))
    For Each Item In colForSearch
        If itemVal = LCase(Trim(Item)) Then
            resBool = True
            Exit For
        End If
    Next Item
    isExistInCol = resBool
End Function

