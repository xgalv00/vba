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
    Unload Me
End Sub

Private Sub chooseSrcBtn_Click()
    'test
    Dim Filt As String
    Dim FilterIndex As Integer
    'Dim FileName As Variant
    Dim Title As String
    Dim i As Integer
    Dim Msg As String
    Dim flw As New FileWorker
    
    ' Set up list of file filters
    Filt = "Excel files (*.xlsx;*.xltx;*.xlsm;*.xltm),*.xlsx;*.xltx;*.xlsm;*.xltm"
    FilterIndex = 1
    ' Set the dialog box caption
    Title = "Выберите файлы для консолидации"
    ' Get the file name
    
    FileName = Application.GetOpenFilename _
                (FileFilter:=Filt, _
                FilterIndex:=FilterIndex, _
                Title:=Title, _
                MultiSelect:=False)
    ' Exit if dialog box canceled
    If FileName = False Then
        MsgBox "Пожалуйста, выберите файл"
        Exit Sub
    End If
    ' Display full path and name of the files
    Msg = CStr(FileName)
    
    copyMineUF.srcNameLbl.Caption = Msg
    copyMineUF.srcNameLbl.ForeColor = vbBlack
    copyMineUF.srcNameLbl.ControlTipText = "Имя файла из которого будет выполнятся копирование"
    copyMineUF.mineManLbl.ForeColor = vbRed
    
    'colCopyUF.execBtn.Enabled = True
End Sub

Private Sub copyBtn_Click()
    Call copyMineFromFile
End Sub

Private Sub mineCmBx_Change()

    copyMineUF.copyBtn.Enabled = True
    copyMineUF.mineLbl.ForeColor = vbBlack
End Sub

Private Sub mineManCmBx_Change()

    copyMineUF.mineManLbl.ForeColor = vbBlack
    copyMineUF.mineLbl.ForeColor = vbRed
End Sub
