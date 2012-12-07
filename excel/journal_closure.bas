Attribute VB_Name = "journal_closure"
Sub UseCanCheckOut()


    Dim xlApp As Excel.Application
    Dim wb As Workbook
    Dim xlFile As String
    
    xlFile = "https://workspaces.dtek.com/it/oisup/ProjectSAP/ChangeManagement/Журнал%20регистрации%20изменений%20в%20проектах%20SAP.xlsm"
    
    'Determine if workbook can be checked out.
    If Workbooks.CanCheckOut(xlFile) = True Then
        Workbooks.CheckOut xlFile
        
        Set xlApp = New Excel.Application
        xlApp.Visible = True
        
        Set wb = xlApp.Workbooks.Open(xlFile, , False)
        
        MsgBox wb.Name & " is checked out to you."
        
        wb.CheckIn (True)
    
    Else
    '
        MsgBox "You are unable to check out this document at this time."
    End If


End Sub

Function isValidVal(inVal As String, modName As String) As Boolean
    Dim tmpArr As Variant
    
    If inVal <> "" Then
    
        If InStr(1, inVal, ".") <> 0 Then
            tmpArr = Split(inVal, ".")
            If IsNumeric(tmpArr(1)) And LCase(Trim(tmpArr(0))) = LCase(Trim(modName)) Then
                isValidVal = True
            End If
        Else
            isValidVal = IsNumeric(inVal)
        End If
    
    End If

End Function

