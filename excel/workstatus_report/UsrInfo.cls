VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim usrName As String
Dim usrLog As String
Public usrType As String
Dim usrEmail As String
Dim compColl As Collection
Dim shtToWork As Worksheet
Dim cachedSht As Worksheet

Function usr_init() As UsrInfo
    Dim usr As New UsrInfo
    Set cachedSht = ActiveSheet
    usrLog = Environ("USERNAME")
    Call find_usr
    Set usr_init = usr
End Function


Private Sub find_usr()
    Dim foundCell As Range
    'examine company's owners
    Sheets("user_table").Select
    Columns("C:C").Select
    Set foundCell = Selection.find(What:=usrLog, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not foundCell Is Nothing Then
        Set shtToWork = ActiveSheet
        usrType = "usr"
        usrName = Cells(foundCell.Row, foundCell.Column - 1).Value
        usrEmail = Cells(foundCell.Row, foundCell.Column + 1).Value
        Set compColl = New Collection
        Do While Not foundCell Is Nothing
            compColl.Add Cells(foundCell.Row, foundCell.Column - 2).Value
            Set foundCell = Selection.FindNext
        Loop
        Exit Sub
    End If
    
    'examine msfo users
    Sheets("msfo_table").Select
    Columns("C:C").Select
    Set foundCell = Selection.find(What:=usrLog, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If Not foundCell Is Nothing Then
        Set shtToWork = ActiveSheet
        usrType = "msfo"
        usrName = Cells(foundCell.Row, foundCell.Column - 1).Value
        usrEmail = Cells(foundCell.Row, foundCell.Column + 1).Value
        Set compColl = New Collection
        Do While Not foundCell Is Nothing
            compColl.Add Cells(foundCell.Row, foundCell.Column - 2).Value
            Set foundCell = Selection.FindNext
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
