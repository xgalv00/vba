VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub refreshBtn_Click()
    rangeValLbl.Caption = ActiveCell.value
End Sub


Private Sub writeBtn_Click()
    Dim rowToIns As Integer
    
    Module1.addRecord
    Unload Me
    Module1.writeReport
End Sub

