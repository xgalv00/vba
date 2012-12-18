VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public modName As String

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim validVal As Boolean
    Dim isFree As Boolean
    Dim targIsClear As Boolean
    
    'targIsClear = False
    If Target.Column = 2 Then
        'if change code added
        modName = Cells(Target.Row, Target.Column + 1).value
        validVal = journal_closure.isValidVal(Target.value, modName)
        
        If validVal Then
            journal_closure.UseCanCheckOut Target.value, modName
        Else
        
            If targIsClear Then
                MsgBox "Value that you have entered is incorrect. Correct format is change number only or module name.change number (e.g. MM.101)"
                Exit Sub
            End If
            targIsClear = True
            Target.value = ""
            
        End If
    
    End If
    
End Sub

