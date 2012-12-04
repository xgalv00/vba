Attribute VB_Name = "journal_cleaner_excel_logging"
    Dim devJour As Workbook, chanJour As Workbook
    Dim devWSht As Worksheet, chanWSht As Worksheet
    Dim resWBook As Workbook
    Dim tmpWSht As Worksheet


Sub journalCleaning()
    
    Dim tmpRow As Integer, tmpCol As Integer, tmpCell As Range, tmpStr As String, foundCell As Range 'helper variables
    Dim clw As New CellWorker, flw As New FileWorker
    Dim devCode As String, chanCode As String, modName As String, developerName As String
    Dim devCodeCol As Integer, chanCodeCol As Integer
    Dim chkModNameToo As Boolean
    Dim tmpArray As Variant
    Dim xmlName As String, desktopPath As String, rootTagName As String
    Dim cachedCell As Range
    Dim cachedSht As Worksheet
    Dim prevVal As String
    Dim firstFoundCell As String
    Dim stopCheck As Boolean
    
    Application.EnableEvents = False
    
    desktopPath = Environ("USERPROFILE")
    
    desktopPath = desktopPath & "\Desktop\"
    
    'logging
    Set resWBook = Workbooks.Add
    resWBook.SaveAs desktopPath & "Результат_обработки_журналов.xlsx", 51
    
    
    
    Set chanJour = Workbooks("Журнал регистрации изменений в проектах SAP.xlsm")
    Set devJour = Workbooks("журнал разработок_new.xlsm")
    Set devWSht = devJour.Sheets(1)
    Set chanWSht = chanJour.Sheets("журнал запросов на измение")
    chanCodeCol = 2
    devCodeCol = 4
    devNameCol = 41
    
    
    Call replRusByEng
    
'$$$
    
    'dev codes cleaning
    
    'logging
    Set tmpWSht = resWBook.Sheets(2)
    tmpWSht.Name = "Ошибки журнала разработок"
    tmpWSht.Activate
    Set tmpCell = ActiveCell
    tmpCell.value = "Тип"
    clw.move_right(tmpCell).value = "Наименование ошибки/изменения"
    clw.move_right(tmpCell, 2).value = "Код разработки"
    clw.move_right(tmpCell, 3).value = "Код изменения"
    clw.move_right(tmpCell, 4).value = "Адрес ячейки кода изменения/разработки"
    clw.move_right(tmpCell, 5).value = "Предыдущее значение кода разработок в журнале изменений"
    clw.move_down(tmpCell).Activate
    
    
    devWSht.Activate
    tmpRow = 3
    chkModNameToo = False
    
    'work
    Set tmpCell = Cells(tmpRow, devCodeCol)
    tmpCell.Select
    Do While devWSht.UsedRange.Rows.CountLarge + 5 > tmpCell.Row
    
        chanCode = Cells(tmpCell.Row, chanCodeCol).value
        modName = Cells(tmpCell.Row, chanCodeCol + 1).value
        devCode = tmpCell.value
        
        If devCode = "" And modName <> "" Then
            'error in dev journal dev code is omitted
            'logging
            Set cachedSht = ActiveSheet
            
            tmpWSht.Activate
            Set cachedCell = ActiveCell
            cachedCell.value = "Ошибка"
            clw.move_right(cachedCell).value = "Пропущен код разработки в журнале разработок"
            'clw.move_right(cachedCell, 2).value = devCode
            'clw.move_right(cachedCell, 3).value = chanCode & "." & modName
            clw.move_right(cachedCell, 4).value = tmpCell.Address(False, False)
            clw.move_down(cachedCell).Activate
            
            cachedSht.Activate
            'work
            Set tmpCell = clw.move_down(tmpCell)
        Else
            
            If chanCode <> "" Then
                chanCode = Trim(chanCode)
                
                '@todo add checks for differrent change code's formats
                'change code cleaning
                If InStr(1, chanCode, modName) <> 0 Then
                    'if change code contains name of module try to find the dot
                    If InStr(1, chanCode, ".") <> 0 Then
                        tmpArray = Split(chanCode, ".")
                        chanCode = tmpArray(1)
                    Else
                        Set cachedSht = ActiveSheet
                        tmpWSht.Activate
                        Set cachedCell = ActiveCell
                        cachedCell.value = "Ошибка"
                        clw.move_right(cachedCell).value = "Что-то не так с номером изменения в журнале разработок по адресу "
                        clw.move_right(cachedCell, 4).value = Cells(tmpCell.Row, chanCodeCol).Address(False, False)
                        cachedSht.Activate
                        stopCheck = True
                        'Debug.Print "Change code " & chanCode & " in address " & Cells(tmpCell.Row, chanCodeCol).Address & " contains wrong code"
                    End If
                End If
                If Not stopCheck Then
                    chanWSht.Activate
                    Columns("B:B").Select
                    Set foundCell = Selection.Find(What:=chanCode, After:=ActiveCell, LookIn:=xlFormulas, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
                    
                    
                    If Not foundCell Is Nothing Then
                        firstFoundCell = foundCell.Address
                        
                        Do While LCase(Trim(modName)) <> LCase(Trim(Cells(foundCell.Row, (devCodeCol - 1)).value))
                            
                            Set foundCell = Selection.FindNext
                            'Set foundCell = Selection.Find(What:=chanCode, After:=ActiveCell, LookIn:=xlFormulas, _
                            'LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                            'MatchCase:=False, SearchFormat:=False)
                            If firstFoundCell = foundCell.Address Then
                                Set foundCell = Nothing
                                firstFoundCell = ""
                                Exit Do
                            End If
                        
                        Loop
                    End If
                    
                    
                    'maybe add here do while foundcell is nothing
                    If foundCell Is Nothing Then
                    
                            Set cachedSht = ActiveSheet
                
                            tmpWSht.Activate
                            Set cachedCell = ActiveCell
                            cachedCell.value = "Ошибка"
                            clw.move_right(cachedCell).value = "Такого номера изменений (или комбинации модуля и номера) нет в журнале изменений, но есть в журнале разработок"
                            clw.move_right(cachedCell, 2).value = devCode
                            clw.move_right(cachedCell, 3).value = chanCode
                            clw.move_right(cachedCell, 4).value = tmpCell.Address(False, False)
                            clw.move_down(cachedCell).Activate
                            
                            cachedSht.Activate
                    
                    Else
                        foundCell.Select
                    
                        If Cells(foundCell.Row, devCodeCol).value = "" Then
                            'logging
                        
                            Set cachedSht = ActiveSheet
                        
                            tmpWSht.Activate
                            Set cachedCell = ActiveCell
                            cachedCell.value = "Изменение"
                            clw.move_right(cachedCell).value = "Был добавлен номер разработки в журнал изменений"
                            clw.move_right(cachedCell, 2).value = devCode
                            clw.move_right(cachedCell, 4).value = Cells(foundCell.Row, devCodeCol).Address(False, False)
                            clw.move_down(cachedCell).Activate
                            cachedSht.Activate
                            
                            Cells(foundCell.Row, devCodeCol).value = devCode
                            
                        Else
                            
                                                                                    'logging
                                prevVal = Cells(foundCell.Row, devCodeCol).value
                                'if previous and new values of dev codes match do nothing
                                If prevVal <> devCode Then
                                    Set cachedSht = ActiveSheet
                                    
                                    tmpWSht.Activate
                                    Set cachedCell = ActiveCell
                                    cachedCell.value = "Изменение"
                                    clw.move_right(cachedCell).value = "Номер разработки в журнале изменений был изменен на"
                                    clw.move_right(cachedCell, 2).value = devCode
                                    'cachedCell.Select
                                    clw.move_right(cachedCell, 4).value = Cells(foundCell.Row, devCodeCol).Address(False, False)
                                    'chanWSht.Activate
                                    'prevVal = Cells(foundCell.Row, devCodeCol).value
                                    'tmpWSht.Activate
                                    clw.move_right(cachedCell, 5).value = prevVal
                                    clw.move_down(cachedCell).Activate
                                    cachedSht.Activate
                                    'cache
                                    
                                    Cells(foundCell.Row, devCodeCol).value = devCode
                                End If
                                prevVal = ""
                        
                        End If 'If Cells(foundCell.Row, devCodeCol).value = "" Then
                        
                    End If 'If foundCell Is Nothing Then
                
                End If 'if not stopCheck then
                
            End If
            
        End If
                
        stopCheck = False
        Set foundCell = Nothing
        devWSht.Activate
        Set tmpCell = clw.move_down(tmpCell)

    Loop
    
    
'$$$
    
    
    
    'change codes cleaning
    'logging
    Set tmpWSht = resWBook.Sheets(3)
    tmpWSht.Name = "Ошибки журнала изменений"
    tmpWSht.Activate
    Set tmpCell = ActiveCell
    tmpCell.value = "Тип"
    clw.move_right(tmpCell).value = "Наименование ошибки/изменения"
    clw.move_right(tmpCell, 2).value = "Код разработки"
    clw.move_right(tmpCell, 3).value = "Код изменения"
    clw.move_right(tmpCell, 4).value = "Адрес ячейки кода изменения/разработки"
    clw.move_right(tmpCell, 5).value = "Предыдущее значение кода изменений в журнале разработок"
    clw.move_down(tmpCell).Activate
    
    'work
    chanWSht.Activate
    tmpRow = 4
    
    Set tmpCell = Cells(tmpRow, chanCodeCol)
    tmpCell.Select
    Do While chanWSht.UsedRange.Rows.CountLarge + 5 > tmpCell.Row
    
        chanCode = tmpCell.value
        devCode = Cells(tmpCell.Row, devCodeCol)
        modName = Cells(tmpCell.Row, chanCodeCol + 1).value
        developerName = Cells(tmpCell.Row, devNameCol).value

        
            If devCode <> "" Then
                devCode = Trim(devCode)
                
                devWSht.Activate
                Columns("D:D").Select
                Set foundCell = Selection.Find(What:=devCode, After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                
                If foundCell Is Nothing Then
                
                    'logging
                    tmpWSht.Activate
                    
                    Set cachedCell = ActiveCell
                    cachedCell.value = "Ошибка"
                    clw.move_right(cachedCell).value = "Код разработки отсутствует в журнале разработок, но есть в журнале изменений"
                    clw.move_right(cachedCell, 2).value = devCode
                    clw.move_right(cachedCell, 3).value = chanCode
                    clw.move_right(cachedCell, 4).value = Cells(tmpCell.Row, devCodeCol).Address(False, False)
                    clw.move_down(cachedCell).Activate
                Else
                
                    prevVal = Cells(foundCell.Row, chanCodeCol).value
                    'if previous and new values of dev codes match do nothing
                    If prevVal <> chanCode Then
                        Set cachedSht = ActiveSheet
                        
                        tmpWSht.Activate
                        Set cachedCell = ActiveCell
                        cachedCell.value = "Изменение"
                        clw.move_right(cachedCell).value = "Номер изменения в журнале разработок был изменен на"
                        clw.move_right(cachedCell, 2).value = devCode
                        clw.move_right(cachedCell, 3).value = chanCode
                        'cachedCell.Select
                        clw.move_right(cachedCell, 4).value = Cells(foundCell.Row, devCodeCol).Address(False, False)
                        'chanWSht.Activate
                        'prevVal = Cells(foundCell.Row, devCodeCol).value
                        'tmpWSht.Activate
                        clw.move_right(cachedCell, 5).value = prevVal
                        clw.move_down(cachedCell).Activate
                        cachedSht.Activate
                        'cache
                        
                        Cells(foundCell.Row, chanCodeCol).value = chanCode
                    End If
                    prevVal = ""
                    
                    'tmpStr = Cells(foundCell.Row, chanCodeCol).Value
                    'If Cells(foundCell.Row, chanCodeCol).value = "" Then
                        'logging
                        'cache
                        
                        'work
                    '    Cells(foundCell.Row, chanCodeCol).value = chanCode
                    'End If
                End If
                
                'foundCell.Select
                Set foundCell = Nothing
            Else
                    
                If developerName <> "" Then
                        'logging
                        'cache
                        Set cachedSht = ActiveSheet
            
                        tmpWSht.Activate
                        Set cachedCell = ActiveCell
                        cachedCell.value = "Ошибка"
                        clw.move_right(cachedCell).value = "Отсутствует номер разработки в журнале изменений"
                        clw.move_right(cachedCell, 3).value = chanCode
                        clw.move_right(cachedCell, 4).value = tmpCell.Address(False, False)
                        'clw.move_right(cachedCell, 4).value =
                        clw.move_down(cachedCell).Activate
                        
                        cachedSht.Activate
                        
                End If
            End If
                
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

    
    'logging
    Set tmpWSht = resWBook.Sheets(1)
    tmpWSht.Name = "Замена_русских_букв"
    tmpWSht.Activate
    Range("A1").value = "Ошибки транслитерации в журнале изменений"
    Set tmpCell = Range("A2")
    tmpCell.value = "Русская буква была заменена"
    Set tmpCell = clw.move_right(tmpCell)
    tmpCell.value = "В ячейке"
    Cells(tmpCell.Row + 1, tmpCell.Column - 1).Activate
    
    'work
    chanWSht.Activate
    Set rangeForClean = Range("B4", Cells(chanWSht.UsedRange.Rows.CountLarge, 4))
    rangeForClean.Select
    Call txtCleaning(rangeForClean)
    
    'logging
    tmpWSht.Activate
    Set tmpCell = ActiveCell
    tmpCell.value = "Ошибки транслитерации в журнале разработок"
    Set tmpCell = clw.move_down(tmpCell)
    tmpCell.value = "Русская буква была заменена"
    Set tmpCell = clw.move_right(tmpCell)
    tmpCell.value = "В ячейке"
    Cells(tmpCell.Row + 1, tmpCell.Column - 1).Activate
    
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
            'logging
            foundCellAddr = foundCell.Address(False, False)
            Set cachedSht = ActiveSheet
            tmpWSht.Activate
            Set tmpCell = ActiveCell
            tmpCell.value = engLetters(i)
            clw.move_right(tmpCell).value = foundCellAddr
            clw.move_down(tmpCell).Activate
            cachedSht.Activate
            Set cachedSht = Nothing
            Set tmpCell = Nothing
            
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

