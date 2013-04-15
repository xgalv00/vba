VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} copyMineUF 
   Caption         =   "Консолидация юрлица"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "copyMineUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "copyMineUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancelBtn_Click()
    Call unloadCopyMineUF
End Sub

Private Sub chooseSrcBtn_Click()
    ' Display full path and name of the files
    Call Controller.proccesFileSelection
    
End Sub

Private Sub copyBtn_Click()
    Call Controller.copyBtnClicked
End Sub

Private Sub mineCmBx_Change()
    Call Controller.mineCmBx_Changed
End Sub

Private Sub mineManCmBx_Change()
    copyMineUF.mineCmBx.RowSource = Controller.computerRowSource("mineCmBx")
    copyMineUF.mineManLbl.ForeColor = vbBlack
    Call Controller.mineManCmBx_Changed
End Sub
