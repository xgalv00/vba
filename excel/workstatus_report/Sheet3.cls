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
Dim oldVal As String

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim mailSndr As New MailSender
    Dim isAuthorized As Boolean
    If Not technicalChange Then 'prevent from calling of record_change procedure during programmatical creation of status' list
        Call initialize_WS_variables
        If isInWorkrange(Target) Then 'record_change procedure will be called only if changed cell lays in appropriate range
            isAuthorized = True
            If isAuthorized Then
                If mailSndr.sendMsg(Target.Address, oldVal) Then
                    If record_change(Target.Address) Then
                        Call mailSndr.completeSendMsg
                        MsgBox "Approval text"
                    End If
                End If
            End If
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
