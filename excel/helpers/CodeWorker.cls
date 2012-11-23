VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub GetGUID(VBProj As VBIDE.VBProject)

   Dim i As Integer
   Dim strResults As String


   For i = 1 To VBProj.References.Count
      strResults = VBProj.References(i). _
      Name & vbTab

      strResults = strResults & VBProj. _
      References(i).GUID & vbTab

      strResults = strResults & VBProj. _
      References(i).Major & vbTab

      strResults = strResults & VBProj. _
      References(i).Minor & vbTab

      Debug.Print strResults
   Next i
End Sub


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

Public Function VBComponentExists(VBCompName As String, Optional VBProj As VBIDE.VBProject = Nothing) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns True or False indicating whether a VBComponent named
    ' VBCompName exists in the VBProject referenced by VBProj. If VBProj
    ' is omitted, the VBProject of the ActiveWorkbook is used.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim VBP As VBIDE.VBProject
    If VBProj Is Nothing Then
        Set VBP = ActiveWorkbook.VBProject
    Else
        Set VBP = VBProj
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
