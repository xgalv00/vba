VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim validVal As Boolean
    Dim isFree As Boolean
    Dim modName As String
    
    If Target.Column = 2 Then
        'if change code added
        modName = Cells(Target.Row, Target.Column + 1).value
        validVal = journal_closure.isValidVal(Target.value, modName)
        
        If validVal Then
            journal_closure.UseCanCheckOut
        Else
            Target.value = ""
        End If
    
    End If
    
End Sub
