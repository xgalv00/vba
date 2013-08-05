VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bulkStatusChangeUF 
   Caption         =   "Изменение статуса формы"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3870
   OleObjectBlob   =   "bulkStatusChangeUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bulkStatusChangeUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub cancelBtn_Click()
    Unload bulkStatusChangeUF
End Sub

Private Sub chStatusBtn_Click()
    Dim statusVal As String
    
    statusVal = bulkStatusChangeUF.statusValCmBx.text
    Call bulkStatusChange(statusVal)
    Unload bulkStatusChangeUF
End Sub

Private Sub statusValCmBx_Change()
    If statusValCmBx.text <> "" Then
        Label1.ForeColor = vbBlack
        chStatusBtn.Enabled = True
    Else
        Label1.ForeColor = vbRed
        chStatusBtn.Enabled = False
    End If
End Sub


Private Sub UserForm_Initialize()
    Label1.ForeColor = vbRed
End Sub
