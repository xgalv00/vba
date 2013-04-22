Attribute VB_Name = "UsrInfo"
Dim usrName As String
Dim usrLog As String
Public usrType As String
Dim usrEmail As String
Dim compColl As Collection
Dim shtToWork As Worksheet
Dim cachedSht As Worksheet

Function usr_init() As Collection
    Dim tmpColl As New Collection
    Set cachedSht = ActiveSheet
    usrLog = Environ("USERNAME")
    Call find_usr
    tmpColl.Add usrLog, "login"
    tmpColl.Add usrName, "name"
    tmpColl.Add usrType, "type"
    tmpColl.Add compColl, "company"
    tmpColl.Add usrEmail, "mail"
    Set usr_init = tmpColl
End Function


Private Sub find_usr()
    Dim foundCell As Range
    Dim nextFoundCell As Range
    Dim foundCellAddr As String
    'examine company's owners
    Sheets("user_table").Select
    Columns("C:C").Select
    Set foundCell = Selection.find(what:=usrLog, after:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not foundCell Is Nothing Then
        Set shtToWork = ActiveSheet
        usrType = "usr"
        usrName = Cells(foundCell.Row, foundCell.Column - 1).Value
        usrEmail = Cells(foundCell.Row, foundCell.Column + 1).Value
        Set compColl = New Collection
        foundCellAddr = foundCell.Address
        Do While Not foundCell Is Nothing
            compColl.Add Cells(foundCell.Row, foundCell.Column - 2).Value
            Set foundCell = Selection.FindNext
            If foundCellAddr = foundCell.Address Then Exit Do
        Loop
        Exit Sub
    End If
    
    'examine msfo users
    Sheets("msfo_table").Select
    Columns("C:C").Select
    Set foundCell = Selection.find(what:=usrLog, after:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If Not foundCell Is Nothing Then
        Set shtToWork = ActiveSheet
        usrType = "msfo"
        usrName = Cells(foundCell.Row, foundCell.Column - 1).Value
        usrEmail = Cells(foundCell.Row, foundCell.Column + 1).Value
        Set compColl = New Collection
        foundCellAddr = foundCell.Address
        Do While Not foundCell Is Nothing
            compColl.Add Cells(foundCell.Row, foundCell.Column - 2).Value
            Set foundCell = Selection.FindNext
            If foundCellAddr = foundCell.Address Then Exit Do
        Loop
        Exit Sub
    Else
        Debug.Assert False
        '"User does not exist in table"
    End If

End Sub
Function isCompanyInUsrCompColl(compName As String) As Boolean
    
    For Each comp In compColl
        If comp = compName Then
            isCompanyInUsrCompColl = True
            Exit Function
        End If
    Next comp
    
End Function

Function isUsrHasApprType(statVal As String) As Boolean

    If (statVal = "Данные содержат ошибки" Or statVal = "Принято" Or statVal = "По умолчанию") And usrType = "msfo" Then
        isUsrHasApprType = True
    ElseIf (statVal = "Данные внесены" Or statVal = "Ввод начат" Or statVal = "По умолчанию") And usrType = "usr" Then
        isUsrHasApprType = True
    End If
End Function


