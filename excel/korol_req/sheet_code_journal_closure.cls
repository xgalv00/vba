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
    Dim tmpArray As Variant
    
    
    'targIsClear = False
    If Target.Column = 2 Then
        'if change code added
        'modName = Cells(Target.Row, Target.Column + 1).value
        'validVal = journal_closure.isValidVal(Target.value)
        
        If Not IsNumeric(Target.value) Then
        '    valForChange = Cells(Target.Row, Target.Column + 2).value
        '    journal_closure.UseCanCheckOut Target.value, modName, valForChange
        'Else
            'If Target.value = "" Then
                'maybe add some reaction on empty value
                'Debug.Print "some"
            'Else
                MsgBox "Value you have entered is incorrect. Correct format is change number only (e.g. 101)"
                Application.EnableEvents = False
                Target.value = ""
                Application.EnableEvents = True
                Target.Activate
            'End If
        End If
        
        'check if target needs to be cleared when non-existent value was entered
        'If targNeedClear Then
        '    Application.EnableEvents = False
        '    Target.value = ""
        '    Application.EnableEvents = True
        '    targNeedClear = False
        'End If
    
    End If
    
    If Target.Column = 4 Then
        'if change code added
        modName = Cells(Target.Row, Target.Column - 1).value
        If modName <> "" Then
            'validVal = journal_closure.isValidDevName(Target.value, modName)
            If Target.value <> "" Then
                'must contain dot.  must contain correct module name.
                'needs cleaning and preparing before validation
                Application.EnableEvents = False
                Target.value = journal_closure.remRusLetters(UCase(Trim(Target.value)))
                Cells(Target.Row, Target.Column - 1).value = journal_closure.remRusLetters(UCase(Trim(modName)))
                Application.EnableEvents = True
                
                'must contain dot
                If InStr(1, Target.value, ".") <> 0 Then
                    tmpArray = Split(Target.value, ".")
                    'letters before dot should be at least part of module name
                    If InStr(1, modName, tmpArray(0)) <> 0 Then
                        'second part should be number
                        If IsNumeric(tmpArray(1)) Then
                            validVal = True
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "You must enter module name first"
            Application.EnableEvents = False
            Target.value = ""
            Application.EnableEvents = True
            Cells(Target.Row, Target.Column - 1).Activate
            Exit Sub
        End If
        
        If Not validVal Then
            MsgBox "Value you have entered is incorrect. Correct format is module.dev number (e.g. MM.101)"
            Application.EnableEvents = False
            Target.value = ""
            Application.EnableEvents = True
            Target.Activate
        End If
        
    End If
End Sub




