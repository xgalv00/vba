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







Public msgWasSent As Boolean

Private Sub CommandButtonCancel_Click()
    
    Unload Me
    
End Sub

Private Sub CommandButtonSend_Click()
    '
    'Dim mailSndr As New MailSender
    msgWasSent = True
    Hide
    '
End Sub


