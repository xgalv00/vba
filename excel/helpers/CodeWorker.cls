VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function ifRefExist(refGUID As String, vbProj As VBIDE.VBProject) As Boolean
    'tests if given reference exists in a given vba project
    Dim i As Integer
    Dim strResults As String
    
    
    For i = 1 To vbProj.References.Count
       If refGUID = vbProj.References(i).GUID Then
            ifRefExist = True
            Exit Function
       End If
    Next i

End Function


Sub CreateEventProcedure(VBComp As VBIDE.VBComponent)
    Dim LineNum As Long
    Const DQUOTE = """" ' one " character
    Dim codeMod As VBIDE.CodeModule
    
    Set codeMod = VBComp.CodeModule

    
    With codeMod
        LineNum = .CreateEventProc("Activate", "Workbook")
        LineNum = LineNum + 1
        'On Error Resume Next
        'Call onWorkBookActivate
        'If Err.Number <> 0 Then
        '    Debug.Print
        'End If
        On Error GoTo 0
        'maybe add here some error checking lines
        '.InsertLines LineNum, "    On error resume next"
        .InsertLines LineNum, "    Call onWorkBookActivate"
    End With
End Sub

Sub DeleteProcedureFromModule(VBComp As VBIDE.VBComponent)
    Dim StartLine As Long
    Dim NumLines As Long
    Dim ProcName As String
    Dim codeMod As VBIDE.CodeModule
    
    Set codeMod = VBComp.CodeModule

    
    ProcName = "DeleteThisProc"
    With codeMod
        StartLine = .ProcStartLine(ProcName, vbext_pk_Proc)
        NumLines = .ProcCountLines(ProcName, vbext_pk_Proc)
        .DeleteLines StartLine:=StartLine, Count:=NumLines
    End With
End Sub

Public Function VBComponentExists(VBCompName As String, Optional vbProj As VBIDE.VBProject = Nothing) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns True or False indicating whether a VBComponent named
    ' VBCompName exists in the VBProject referenced by VBProj. If VBProj
    ' is omitted, the VBProject of the ActiveWorkbook is used.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim VBP As VBIDE.VBProject
    If vbProj Is Nothing Then
        Set VBP = ActiveWorkbook.VBProject
    Else
        Set VBP = vbProj
    End If
    On Error Resume Next
    VBComponentExists = CBool(Len(VBP.VBComponents(VBCompName).Name))

End Function

Public Function ListProcedures(VBComp As VBIDE.VBComponent) As Collection
        Dim LineNum As Long
        Dim NumLines As Long
        Dim ProcName As String
        Dim ProcKind As VBIDE.vbext_ProcKind
        Dim tmpStr As String
        Dim tmpColl As New Collection
        Dim codeMod As VBIDE.CodeModule
        
        Set codeMod = VBComp.CodeModule
        
        With codeMod
            
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                ProcName = .ProcOfLine(LineNum, ProcKind)
                tmpColl.Add ProcName
                tmpStr = ProcKindString(ProcKind)
                LineNum = .ProcStartLine(ProcName, ProcKind) + _
                        .ProcCountLines(ProcName, ProcKind) + 1
            Loop
        End With
        If tmpColl.Count > 0 Then
            Set ListProcedures = tmpColl
        End If
    End Function
    
    
Private Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
    Select Case ProcKind
        Case vbext_pk_Get
            ProcKindString = "Property Get"
        Case vbext_pk_Let
            ProcKindString = "Property Let"
        Case vbext_pk_Set
            ProcKindString = "Property Set"
        Case vbext_pk_Proc
            ProcKindString = "Sub Or Function"
        Case Else
            ProcKindString = "Unknown Type: " & CStr(ProcKind)
    End Select
End Function

Sub ListModules()
    Dim vbProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim WS As Worksheet
    Dim Rng As Range
    
    Set vbProj = ActiveWorkbook.VBProject
    Set WS = ActiveWorkbook.Worksheets("Sheet1")
    Set Rng = WS.Range("A1")
    
    For Each VBComp In vbProj.VBComponents
        Rng(1, 1).value = VBComp.Name
        Rng(1, 2).value = ComponentTypeToString(VBComp.Type)
        Set Rng = Rng(2, 1)
    Next VBComp
End Sub


Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function

Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
            FolderName As String, _
            Optional FileName As String, _
            Optional OverwriteExisting As Boolean = True) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Extension As String
Dim FName As String
Extension = GetFileExtension(VBComp:=VBComp)
If Trim(FileName) = vbNullString Then
    FName = VBComp.Name & Extension
Else
    FName = FileName
    If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
        FName = FName & Extension
    End If
End If

If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
    FName = FolderName & FName
Else
    FName = FolderName & "\" & FName
End If

If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
        Kill FName
    Else
        ExportVBComponent = False
        Exit Function
    End If
End If

VBComp.Export FileName:=FName
ExportVBComponent = True

End Function

Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function
