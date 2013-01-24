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
Dim curWSht As Worksheet, tmpWSht As Worksheet
Dim curForm As String, folForWrite As String
Public flagEnd As Boolean


Private Sub refreshBtn_Click()
    rangeValLbl.Caption = ActiveCell.value
End Sub

Public Sub setBooks(fName As String, folName As String)
    Dim curWB As Workbook, tmpWB As Workbook
    
    curForm = fName
    folForWrite = folName
    Set curWB = ActiveWorkbook
    Application.EnableEvents = False
    Set tmpWB = Workbooks.Open(folForWrite & curForm)
    'tmpWB.Unprotect "sap"
    Set curWSht = curWB.Sheets(1)
    Set tmpWSht = tmpWB.Sheets(1)
    tmpWSht.Activate
    
End Sub


Private Sub writeBtn_Click()
    Dim rowToIns As Integer
    
    rowToIns = curWSht.UsedRange.Rows.Count + curWSht.UsedRange.Row
    curWSht.Cells(rowToIns, 1).value = folForWrite
    curWSht.Cells(rowToIns, 2).value = curForm
    curWSht.Cells(rowToIns, 3).value = rangeValLbl.Caption
    ActiveWorkbook.Close False
    Application.EnableEvents = True
    flagEnd = True
    'Unload Me
End Sub
