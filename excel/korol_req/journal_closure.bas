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
        

            Cells(foundCell.Row, foundCell.Column + 2).Value = valForChange
            
            'check if different names of module used, only check do nothing
            If modName <> Trim(Cells(foundCell.Row, foundCell.Column + 1).Value) Then
                errMsg = "Возможная ошибка " & vbCrLf
                errMsg = errMsg & "Names of modules don't match, maybe this is not a mistake but check it please" & vbCrLf
                errMsg = errMsg & "Module's name from dev journal " & modName & "; Module's name from change journal " & Trim(Cells(foundCell.Row, foundCell.Column + 1).Value)
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

