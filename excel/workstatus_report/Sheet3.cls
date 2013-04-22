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
    

    If Not technicalChange Then 'prevent from calling of record_change procedure during programmatical creation of status' list
        Call initialize_WS_variables
        If isInWorkrange(Target) Then 'record_change procedure will be called only if changed cell lays in appropriate range
            If Target.Count > 1 Then
                MsgBox "����� �������� ������ ������ � ����� ������"
                Exit Sub
            End If
            Call unhide_everything
            Set changer = usr_init()
            Debug.Assert Not changer Is Nothing
            If isAuthorized(Target) Then
                '@todo add check that value is one range
                If mailSndr.sendMsg(Target.Address, CStr(oldVal), changer("name")) Then
                    If record_change(Target.Address) Then
                        Call mailSndr.completeSendMsg
                        MsgBox "Approval text"
                    End If
                End If
            Else
                Debug.Assert False
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