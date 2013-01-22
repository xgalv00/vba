Attribute VB_Name = "journal_closure"
Sub UseCanCheckOut(targetVal As String, modName As String, valForChange As String)


    Dim xlApp As Excel.Application
    Dim wb As Workbook
    Dim xlFile As String
    Dim foundCell As Range
    Dim wSht As Worksheet
    Dim errMsg As String
    
    'xlFile = "https://workspaces.dtek.com/it/oisup/ProjectSAP/ChangeManagement/Журнал%20регистрации%20изменений%20в%20проектах%20SAP.xlsm"
    xlFile = "https://workspaces.dtek.com/it/oisup/ProjectSAP/ChangeManagement/test.xlsm"
    
    'prepare values
    targetVal = Trim(targetVal)
    modName = Trim(modName)
    

    'many events in workbook to open
    Application.EnableEvents = False
    
    'Determine if workbook can be checked out.
    If Workbooks.CanCheckOut(xlFile) = True Then
        

        Workbooks.CheckOut xlFile
        
        Set wb = Workbooks.Open(xlFile, , False)
        Set wSht = wb.Sheets("журнал запросов на измение")
        wSht.Select
        
        wSht.Columns("B:B").Select
    
        Set foundCell = Selection.Find(What:=targetVal, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
        If foundCell Is Nothing Then
            MsgBox "Change number that you have entered doesn't exist"
            '@todo clear  value that have been entered in developers journal
            Лист1.targNeedClear = True
        Else
        

            Cells(foundCell.Row, foundCell.Column + 2).value = valForChange
            
            'check if different names of module used, only check do nothing
            If modName <> Trim(Cells(foundCell.Row, foundCell.Column + 1).value) Then
                errMsg = "Возможная ошибка " & vbCrLf
                errMsg = errMsg & "Names of modules don't match, maybe this is not a mistake but check it please" & vbCrLf
                errMsg = errMsg & "Module's name from dev journal " & modName & "; Module's name from change journal " & Trim(Cells(foundCell.Row, foundCell.Column + 1).value)
            End If
            
            MsgBox valForChange & " было вставлено в " & "[" & wb.Name & "]!" & wSht.Name & "." & Cells(foundCell.Row, foundCell.Column + 2).Address(False, False) & vbCrLf & vbCrLf & vbCrLf & errMsg
            errMsg = ""

        End If
        
        Application.EnableEvents = True
        wb.CheckIn (True)
        
        
        
    Else
    '
        MsgBox "You are unable to check out this document at this time. Please try again later."
    End If


End Sub

Function isValidVal(inVal As String) As Boolean
    
    If inVal <> "" Then
        isValidVal = IsNumeric(inVal)
    End If

End Function

Function isValidDevName(inVal As String, modName As String) As Boolean
    Dim tmpArray As Variant

    If inVal <> "" Then
        'must contain dot.  must contain correct module name.
        'needs cleaning and preparing before validation
        inVal = remRusLetters(UCase(Trim(inVal)))
        modName = remRusLetters(UCase(Trim(modName)))
        
        'must contain dot
        If InStr(0, inVal, ".") <> 0 Then
            tmpArray = Split(inVal, ".")
            If InStr(0, modName, tmpArray(0)) <> 0 Then
                isValidDevName = True
            End If
        End If
        
    End If

End Function

Function remRusLetters(valForClean As String) As String
    
    Dim i As Integer, rusLetters As Variant, engLetters As Variant
    Dim cleanedVal As String
    
    cleanedVal = valForClean
    rusLetters = Array("А", "В", "С", "Е", "Н", "К", "М", "О", "Р", "Т", "Х", "У")
    engLetters = Array("A", "B", "C", "E", "H", "K", "M", "O", "P", "T", "X", "Y")
    
    For i = 0 To UBound(rusLetters)
    
        If InStr(0, valForClean, rusLetters(i)) <> 0 Then
            cleanedVal = Replace(valForClean, rusLetters(i), engLetters(i))
        End If
    
    Next i

    remRusLetters = cleanedVal
End Function

