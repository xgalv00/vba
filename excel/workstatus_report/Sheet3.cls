VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim oldVal As Variant

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim mailSndr As New MailSender
    Dim changer As Collection
    Dim isAauthString As String
    

    If Not technicalChange Then 'prevent from calling of record_change procedure during programmatical creation of status' list
        Call initialize_WS_variables
        If isInWorkrange(Target) Then 'record_change procedure will be called only if changed cell lays in appropriate range
            If Target.Count > 1 Then
                MsgBox "Можно изменять статус только в одной ячейке"
                Exit Sub
            End If
            Call unhide_everything
            Set changer = usr_init()
            Debug.Assert Not changer Is Nothing
            isAauthString = isAuthorized(Target)
            If isAauthString = "ok" Then
                If mailSndr.sendMsg(Target.Address, CStr(oldVal), changer("name")) Then
                    If record_change(Target.Address) Then
                        If mailSndr.completeSendMsg = "ok" Then
                            MsgBox "Статус был изменен, и если этому статусу необходима отправка уведомления, то оно было отправлено."
                        Else
                            MsgBox "Статус был изменен, но сообщение не было отправлено, потому что Outlook не установлен или недоступен."
                        End If
                        Cells(Target.Row - 1, Target.Column).Select
                    
                    Else
                        MsgBox "Статус не может быть изменен, потому что у вас не хватает полномочий (record_change check)"
                        Call revertChange(Target)
                    End If
                Else
                    MsgBox "Статус не был изменен, сообщение не отправлено"
                    Call revertChange(Target)
                End If
            Else
                'Debug.Assert False
                MsgBox "Статус не может быть изменен, потому что " & isAauthString
                Call revertChange(Target)
                'user not authorized
            End If
            Call hide_everything
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not technicalChange Then 'prevent from calling of record_change procedure during programmatical creation of status' list
        Call initialize_WS_variables
        If isInWorkrange(Target) Then 'if changed cell lays in appropriate range
            oldVal = Target.Value
        End If
    End If
End Sub

Private Sub revertChange(ByVal Target As Range)

    Debug.Assert oldVal <> ""
    Target.Parent.Activate
    Target.Value = oldVal
    
    Range("K10").Select
    
End Sub
