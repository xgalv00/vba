VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bulkStatusChangeUF 
   Caption         =   "UserForm1"
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



Private Sub chStatusBtn_Click()
    Dim statusVal As String
    
    statusVal = bulkStatusChangeUF.statusValCmBx.text
    Call bulkStatusChange(statusVal)
End Sub
