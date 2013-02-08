Attribute VB_Name = "journal_cleaner_excel_logging"
    Dim devJour As Workbook, chanJour As Workbook
    Dim devWSht As Worksheet, chanWSht As Worksheet
    Dim devJournName As String, chanJournName As String
    Dim resWBook As Workbook
    Dim tmpWSht As Worksheet


Sub journalCleaning()
    
    Dim tmpRow As Integer, tmpCol As Integer, tmpCell As Range, tmpStr As String, foundCell As Range 'helper variables
    Dim clw As New CellWorker, flw As New FileWorker
    Dim devCodeChan As String, chanCodeChan As String, modNameChan As String, developerName As String 'values from change journal
    Dim devCodeDev As String, chanCodeDev As String, modNameDev As String   'values from dev journal
    
    Dim chanCodeCol As Integer, devCodeCol As Integer, modNameCol As Integer
    Dim tmpArray As Variant
    Dim xmlName As String, desktopPath As String, rootTagName As String
    Dim prevVal As String
    Dim firstFoundCell As String
    Dim changeColor As Long
    
    
    Application.EnableEvents = False
    
    
    devJournName = "журнал разработок.xlsm"
    chanJournName = "Журнал регистрации изменений в проектах SAP.xlsm"
    
    
    
    On Error Resume Next
    Set chanJour = Workbooks(chanJournName)
    If Err.Number <> 0 Then
        MsgBox "You must open and check out change journal first"
        Exit Sub
    End If
    On Error GoTo 0
    Set devJour = Workbooks(devJournName)
    Set devWSht = devJour.Sheets(1)
    Set chanWSht = chanJour.Sheets("журнал запросов на измение")
    chanCodeCol = 2
    modNameCol = 3
    devCodeCol = 4
    devNameCol = 41
    changeColor = 65535
    
    
    Call replRusByEng
    'excludes defective rows from processing by changing row color
    Call excludeDefects
    
'$$$
    
    'dev codes cleaning
    
    
    
    devWSht.Activate
    tmpRow = 3
    
    'work
    Set tmpCell = Cells(tmpRow, devCodeCol)
    tmpCell.Select
    Do While devWSht.UsedRange.Rows.CountLarge + 5 > tmpCell.Row
        
        'skip rows that are not well-formed
        If tmpCell.Interior.Color <> 16776960 Or Cells(tmpCell.Row, 1).Interior.Color <> 16776960 Then
            chanCodeDev = Trim(Cells(tmpCell.Row, chanCodeCol).value)
            modNameDev = UCase(Trim(Cells(tmpCell.Row, modNameCol).value))
            devCodeDev = UCase(Trim(tmpCell.value))
            
            If chanCodeDev <> "" Then
                
                If tmpCell.Interior.Color = changeColor Or Cells(tmpCell.Row, chanCodeCol).Interior.Color = changeColor Then
                    Call fixColors(tmpCell)
                    Call fixColors(Cells(tmpCell.Row, chanCodeCol))
                End If
                chanWSht.Activate
                Columns("B:B").Select
                Set foundCell = Selection.Find(What:=chanCodeDev, After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                
                
                'exclude repeated modules while
                If Not foundCell Is Nothing Then
                    'foundCell.Activate
                    firstFoundCell = foundCell.Address
                    modNameChan = UCase(Trim(Cells(foundCell.Row, modNameCol).value))
                    'Do While modNameDev <> modNameChan
                    Do While InStr(1, modNameDev, modNameChan) = 0 And InStr(1, modNameChan, modNameDev) = 0
                        'foundCell.Select
                        'Columns("B:B").Select

                        Set foundCell = Selection.FindNext(foundCell)
                        If firstFoundCell = foundCell.Address Then
                            Set foundCell = Nothing
                            firstFoundCell = ""
                            Exit Do
                        End If
                        modNameChan = UCase(Trim(Cells(foundCell.Row, modNameCol).value))
                    Loop
                End If
                
                
                'maybe add here do while foundcell is nothing
                If foundCell Is Nothing Then
                    'return to dev journal
                    devWSht.Activate
                    Call logError(Cells(tmpCell.Row, chanCodeCol), "Такого номера изменений(или комбинации модуля и номера) нет в журнале изменений", True, changeColor)
                    chanWSht.Activate
                
                Else
                    foundCell.Select
                    If foundCell.Interior.Color <> 16776960 Or Cells(foundCell.Row, 1).Interior.Color <> 16776960 Then
                        
                        chanCodeChan = Trim(Cells(foundCell.Row, chanCodeCol).value)
                        modNameChan = UCase(Trim(Cells(foundCell.Row, modNameCol).value))
                        devCodeChan = UCase(Trim(Cells(foundCell.Row, devCodeCol).value))
                        
                        If devCodeChan = "" Then
                        
                            chanWSht.Activate
                            Call logError(Cells(foundCell.Row, devCodeCol), "Был добавлен номер разработки", True, changeColor)
                            Cells(foundCell.Row, devCodeCol).value = devCodeDev
                            
                        Else
                            
                            prevVal = devCodeChan
                            'if previous and new values of dev codes match do nothing
                            
                            If InStr(1, prevVal, devCodeDev) = 0 Then 'this condition is used because of ; is dev codes
                                
                                If Cells(foundCell.Row, devCodeCol).Comment Is Nothing Then
                                    Cells(foundCell.Row, devCodeCol).value = devCodeDev
                                    Call logError(Cells(foundCell.Row, devCodeCol), "Номер разработки был изменен. Предыдущее значение " & prevVal, True, changeColor)
                                Else
                                    Cells(foundCell.Row, devCodeCol).value = prevVal & ";" & devCodeDev
                                    Call logErrWithCom(Cells(foundCell.Row, devCodeCol), "Номер разработки был изменен. Предыдущее значение " & prevVal)
                                End If

                            End If
                            prevVal = ""
                        End If 'devCodeChan = "" Then
                    End If 'If foundCell.Interior.Color <> 16776960 Then
                    
                End If 'If foundCell Is Nothing Then
            
            End If 'if chanCodeDev <> "" then
            
        End If 'If tmpCell.Interior.Color <> 16776960 Then
        
        'cleaning
        chanCodeDev = ""
        modNameDev = ""
        devCodeDev = ""
        chanCodeChan = ""
        modNameChan = ""
        devCodeChan = ""
        Set foundCell = Nothing
        devWSht.Activate
        Set tmpCell = clw.move_down(tmpCell)

    Loop
    
    
'$$$
    
    
    
    'change codes cleaning
    
    'work
    chanWSht.Activate
    tmpRow = 4
    
    Set tmpCell = Cells(tmpRow, chanCodeCol)
    tmpCell.Select
    Do While chanWSht.UsedRange.Rows.CountLarge + 5 > tmpCell.Row
    
        chanCodeChan = Trim(Cells(tmpCell.Row, chanCodeCol).value)
        modNameChan = UCase(Trim(Cells(tmpCell.Row, modNameCol).value))
        devCodeChan = UCase(Trim(Cells(tmpCell.Row, devCodeCol).value))
        developerName = Cells(tmpCell.Row, devNameCol).value
        
        'skip rows that are not well-formed
        If tmpCell.Interior.Color <> 16776960 Or Cells(tmpCell.Row, 1).Interior.Color <> 16776960 Then
            
            If Cells(tmpCell.Row, 1).Interior.Color = changeColor Then
                Call fixColors(Cells(tmpCell.Row, 1))
            End If
            
            If devCodeChan <> "" Then
                
                devWSht.Activate
                Columns("D:D").Select
                Set foundCell = Selection.Find(What:=devCodeChan, After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                
                If foundCell Is Nothing Then
                    'exclude compound dev codes
                    If InStr(1, devCodeChan, ";") = 0 Then
                        chanWSht.Activate
                        Call logError(Cells(tmpCell.Row, devCodeCol), "Код разработки отсутствует в журнале разработок", True, changeColor)
                    End If
                    
                Else
                    chanCodeDev = Trim(Cells(foundCell.Row, chanCodeCol).value)
                    modNameDev = UCase(Trim(Cells(foundCell.Row, modNameCol).value))
                    devCodeDev = UCase(Trim(Cells(foundCell.Row, devCodeCol).value))
                    
                    'check if found cell is not excluded
                    If foundCell.Interior.Color <> 16776960 Or Cells(foundCell.Row, 1).Interior.Color <> 16776960 Then
                        'If Not chanCodeDev = "" Then
                        '    prevVal = chanCodeDev
                            'if previous and new values of dev codes match do nothing
                        '    If prevVal <> chanCodeChan Then
                            
                        '        Call logError(Cells(foundCell.Row, chanCodeCol), "Номер изменения был изменен. Предыдущее значение " & prevVal, True, changeColor)
                        '        Cells(foundCell.Row, chanCodeCol).value = chanCodeChan
                                
                        '    End If
                        '    prevVal = ""
                        'Else
                        If chanCodeDev = "" Then
                        
                            Call logError(Cells(foundCell.Row, chanCodeCol), "Был добавлен номер изменения", True, changeColor)
                            Cells(foundCell.Row, chanCodeCol).value = chanCodeChan
                            
                        End If 'If Not chanCodeDev = "" Then
                        
                    End If 'If foundCell.Interior.Color <> 16776960 Then
                    
                End If 'If foundCell Is Nothing Then
                
                Set foundCell = Nothing
            Else
            
                'if dev code is omitted but developer name is present this is an error
                If developerName <> "" Then
                    chanWSht.Activate
                    Call logError(tmpCell, "Отсутствует номер разработки в журнале изменений", True, changeColor)
                        
                End If
                
            End If 'If devCode <> "" Then
            
        End If 'If tmpCell.Interior.Color <> 16776960 Then
                
        chanCodeChan = ""
        modNameChan = ""
        devCodeChan = ""
        developerName = ""
        chanCodeDev = ""
        modNameDev = ""
        devCodeDev = ""
                
        chanWSht.Activate
        Set tmpCell = clw.move_down(tmpCell)
    Loop
    
    Application.EnableEvents = True
    
End Sub

Private Sub replRusByEng()

    'replace russian letters by english
     Dim tmpRow As Integer, tmpCol As Integer, tmpCell As Range, tmpStr As String, foundCell As Range 'helper variables
    Dim rangeForClean As Range
    Dim clw As New CellWorker, flw As New FileWorker
    
    'work
    chanWSht.Activate
    Set rangeForClean = Range("B4", Cells(chanWSht.UsedRange.Rows.CountLarge, 4))
    rangeForClean.Select
    Call txtCleaning(rangeForClean)
    
    'work
    devWSht.Activate
    Set rangeForClean = Range("B3", Cells(devWSht.UsedRange.Rows.CountLarge, 4))
    rangeForClean.Select
    Call txtCleaning(rangeForClean)
    
End Sub

Private Sub txtCleaning(rangeForClean As Range)
    'find and replace Russian letters by English in a given range
    Dim i As Integer, rusLetters As Variant, engLetters As Variant
    Dim foundCell As Range
    Dim foundCellAddr As Variant
    Dim tmpCell As Range
    Dim cachedSht As Worksheet
    Dim clw As New CellWorker

    
    rusLetters = Array("А", "В", "С", "Е", "Н", "К", "М", "О", "Р", "Т", "Х", "У")
    engLetters = Array("A", "B", "C", "E", "H", "K", "M", "O", "P", "T", "X", "Y")
    'i = 0
    rangeForClean.Select
    For i = 0 To UBound(rusLetters)
        Set foundCell = Selection.Find(What:=rusLetters(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False)
        Do While Not foundCell Is Nothing
            
            'work
            foundCell.Replace What:=rusLetters(i), Replacement:=engLetters(i), LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
            
            foundCell.Activate
            Set foundCell = Nothing
            Set foundCell = Selection.FindNext(After:=ActiveCell)
        Loop
    Next i

End Sub

Sub excludeDefects()
    '@todo refactor this function
    '@todo add color back function
    Dim tmpRow As Integer, chanCodeCol As Integer, devCodeCol As Integer, modNameCol As Integer
    Dim clw As New CellWorker
    Dim tmpCell As Range
    Dim devcode As String, chanCode As String, modName As String
    Dim validVal As Boolean
    Dim inRow As Integer
    
    chanCodeCol = 2
    modNameCol = 3
    devCodeCol = 4
    devNameCol = 41
    
    devWSht.Activate
    tmpRow = 3
    Set tmpCell = Cells(tmpRow, devCodeCol)
    
    Do While tmpCell.value <> "" Or Cells(tmpCell.Row, modNameCol).value <> "" Or Cells(tmpCell.Row + 1, modNameCol).value <> ""
        'cleans previous exclusion
        If Cells(tmpCell.Row, 1).Interior.Color = 16776960 Then
            Call fixRowColor(tmpCell.Row)
        End If
        
        If tmpCell.value = "" Then
            'add comment and highlight this row
            Call logError(tmpCell, "Отсутствует номер разработки")
        Else
            'check format of dev code
            devcode = UCase(Trim(tmpCell.value))
            modName = UCase(Trim(Cells(tmpCell.Row, modNameCol).value))
            validVal = validateSimpleDevCode(devcode, modName) 'maybe replace it by validateDevCodes
            
            If Not validVal Then
                'add comment and highlight this row
                Call logError(tmpCell, "Некорректный формат. Правильный формат модуль.номер разработки (например ММ.101)")
            End If
            
            chanCode = Trim(Cells(tmpCell.Row, chanCodeCol).value)
            'check format of change code
            'change code can be omitted here
            If Not chanCode = "" Then
                If Not IsNumeric(chanCode) Then
                    'add comment and highlight this row
                    Call logError(Cells(tmpCell.Row, chanCodeCol), "Некорректный формат. Правильный формат - номер изменения (например 100)")
                End If
            End If
        End If
        Set tmpCell = clw.move_down(tmpCell)
    Loop
    
    chanWSht.Activate
    tmpRow = 4
    Set tmpCell = Cells(tmpRow, chanCodeCol)
    
    
    Do While tmpCell.value <> "" Or Cells(tmpCell.Row, modNameCol).value <> ""
        'clean previous exclusion
        If Cells(tmpCell.Row, 1).Interior.Color = 16776960 Then
            Call fixRowColor(tmpCell.Row)
        End If
        
        If tmpCell.value = "" Then
            'add comment and highlight this row
            Call logError(tmpCell, "Отсутствует номер изменения")
        Else
            'check format of dev code
            
            devcode = UCase(Trim(Cells(tmpCell.Row, devCodeCol).value))
            modName = UCase(Trim(Cells(tmpCell.Row, modNameCol).value))
            'in this iteration dev code can be omitted
            If Not devcode = "" Then
                validVal = validateDevCodes(devcode, modName)
            
                If Not validVal Then
                    'add comment and highlight this row
                    Call logError(Cells(tmpCell.Row, devCodeCol), "Некорректный формат. Правильный формат модуль.номер разработки (например ММ.101)")
                End If
            End If
            
            chanCode = Trim(Cells(tmpCell.Row, chanCodeCol).value)
            'check format of change code
            If Not IsNumeric(chanCode) Then
                'add comment and highlight this row
                Call logError(tmpCell, "Некорректный формат. Правильный формат - номер изменения (например 100)")
            End If
        End If
        Set tmpCell = clw.move_down(tmpCell)
    Loop

End Sub
Sub logErrWithCom(inCell As Range, comErr As String)

    inCell.Comment.Text Text:=comErr

End Sub

Sub logError(inCell As Range, comErr As String, Optional colorCellOnly As Boolean, Optional fillColor As Long)
    'function for logging
    
    If fillColor = 0 Then
        fillColor = 16776960
    End If
    
    inCell.AddComment
    inCell.Comment.Visible = False
    inCell.Comment.Text Text:=comErr
    If colorCellOnly Then
        'add here format cleaning
        If inCell.Parent.Parent.Name = chanJournName Then
            Cells(inCell.Row, 1).Interior.Color = fillColor 'in change journal color only first cell in a row
        Else
            inCell.ClearFormats
            inCell.Interior.Color = fillColor
        End If

    Else
        Rows(inCell.Row).Select
        Selection.Interior.Color = fillColor
    End If
    
End Sub

Sub fixColors(inCell As Range)
    'makes cell format like nearby cell
    If inCell.Parent.Parent.Name = chanJournName Then
        inCell.ClearFormats
        If Cells(inCell.Row, 2).Comment Is Nothing Then
            Cells(inCell.Row, 2).Comment.Delete
        ElseIf Cells(inCell.Row, 4).Comment Is Nothing Then
            Cells(inCell.Row, 4).Comment.Delete
        End If

    Else
        Cells(inCell.Row, inCell.Column + 1).Select
        Selection.Copy
        inCell.Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
    If Not inCell.Comment Is Nothing Then
       inCell.Comment.Delete
    End If
End Sub
Function validateCompoundDevCode(devcode As String, modName As String) As Boolean
    'validates dev codes that contains ;. it means their relation to change code is n to 1
    Dim i As Integer
    Dim isValidCode As Boolean
    Dim tmpArr As Variant
    
    If InStr(1, devcode, ";") <> 0 Then
        
        tmpArr = Split(devcode, ";")
        For i = 0 To UBound(tmpArr)
            isValidCode = False
            isValidCode = validateSimpleDevCode(Trim(tmpArr(i)), modName)
            If Not isValidCode Then
                Exit For
            End If
        Next i
        
    End If
    
    validateCompoundDevCode = isValidCode

End Function

Function validateSimpleDevCode(devcode As String, modName As String) As Boolean
    Dim tmpArray As Variant
    Dim validVal As Boolean
    
    If InStr(1, devcode, ".") <> 0 Then
        tmpArray = Split(devcode, ".")
        'letters before dot should be at least part of module name
        If InStr(1, modName, tmpArray(0)) <> 0 Then
            'second part should be number
            If IsNumeric(tmpArray(1)) Then
                validVal = True
            End If
        End If
    End If
    
    validateSimpleDevCode = validVal
End Function

Function validateDevCodes(devcode As String, modName As String) As Boolean

    Dim isValidCode As Boolean
    
    If InStr(1, devcode, ";") <> 0 Then
        isValidCode = validateCompoundDevCode(devcode, modName)
    Else
        isValidCode = validateSimpleDevCode(devcode, modName)
    End If
    
    validateDevCodes = isValidCode
    
End Function

Sub fixRowColor(inRow As Integer)
    Dim i As Integer
    'makes row format like row from below
    Rows(inRow + 1).Select
    i = 2
    Do While Selection.Interior.Color = 16776960
        Rows(inRow + i).Select
        i = i + 1
    Loop
    
    Selection.Copy
    Rows(inRow).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    If Not Cells(inRow, 2).Comment Is Nothing Then
       Cells(inRow, 2).Comment.Delete
    ElseIf Not Cells(inRow, 4).Comment Is Nothing Then
        Cells(inRow, 4).Comment.Delete
    End If

End Sub
'Private Sub Workbook_Open()
' Системное событие - сразу после открытия книги
' Загрузить с сервера набор макросов
'  Dim c As Object
'  For Each c In VBProject.VBComponents
'    If Left(c.Name, 9) = "MasterBPC" Then VBProject.VBComponents.Remove c
'  Next c
'  VBProject.VBComponents.Import "\\v-sap-pcts\Communic\Macros\MasterBPC.bas"
'End Sub

'You can add the reference programmatically with code like:
'
'    ThisWorkbook.VBProject.References.AddFromGuid _
'        GUID:="{0002E157-0000-0000-C000-000000000046}", _
'        Major:=5, Minor:=3
'
'
'    Dim vbProj As VBIDE.VBProject
'
'    Set vbProj = ActiveWorkbook.VBProject

