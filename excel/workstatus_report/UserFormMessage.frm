VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMessage 
   Caption         =   "Отправка уведомлений"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   OleObjectBlob   =   "UserFormMessage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public tempSheet As Worksheet

Private Sub CommandButtonCancel_Click()
    
    Unload Me
    
End Sub

Private Sub CommandButtonSend_Click()
    '
    If LabelFileMessage.Caption = "" Then
        SendStatusMail TextBoxTo.Value, TextBoxText.Value, TextBoxSubject.Value
    Else
        SendStatusMailWithSheet TextBoxTo.Value, TextBoxText.Value, TextBoxSubject.Value, tempSheet
    End If
    '
    Hide
    '
End Sub

Public Sub SendWithOutCheck()

    CommandButtonSend_Click
    
End Sub

Private Sub LabelFileMessage_Click()

End Sub

Private Sub UserForm_Click()

End Sub
