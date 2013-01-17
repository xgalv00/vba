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
Public targNeedClear As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim validVal As Boolean
    Dim isFree As Boolean
    Dim valForChange As String
    
    
    'targIsClear = False
    If Target.Column = 2 Then
        'if change code added
        modName = Cells(Target.Row, Target.Column + 1).Value
        validVal = journal_closure.isValidVal(Target.Value)
        
        If validVal Then
            valForChange = Cells(Target.Row, Target.Column + 2).Value
            journal_closure.UseCanCheckOut Target.Value, modName, valForChange
        Else
            If Target.Value = "" Then
                'maybe add some reaction on empty value
                'Debug.Print "some"
            Else
                MsgBox "Value that you have entered is incorrect. Correct format is change number only (e.g. 101)"
                Application.EnableEvents = False
                Target.Value = ""
                Application.EnableEvents = True
            End If


            
        End If
        
        'check if target needs to be cleared when non-existent value was entered
        If targNeedClear Then
            Application.EnableEvents = False
            Target.Value = ""
            Application.EnableEvents = True
            targNeedClear = False
        End If
    
    End If
    
End Sub



