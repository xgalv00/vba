Attribute VB_Name = "Module1"
Dim curWSht As Worksheet, tmpWSht As Worksheet
Dim curForm As String, folForWrite As String

Sub writeReport()
    Dim folPath As String
    Dim fName As String
    
    fName = Dir(folPath)
    writeReportName fName, folPath

End Sub

Sub writeReportName(fName As String, folName As String)
    
    setBooks fName, folName
    UserForm1.Show False
    
End Sub

Public Sub setBooks(fName As String, folName As String)
    Dim curWB As Workbook, tmpWB As Workbook
    
    curForm = fName
    folForWrite = folName
    Set curWB = ActiveWorkbook
    Application.EnableEvents = False
    Set tmpWB = Workbooks.Open(folForWrite & curForm)
    Set curWSht = curWB.Sheets(1)
    Set tmpWSht = tmpWB.Sheets(1)
    tmpWSht.Activate
    
End Sub

Public Sub addRecord()

    rowToIns = curWSht.UsedRange.Rows.Count + curWSht.UsedRange.Row
    curWSht.Cells(rowToIns, 1).value = folForWrite
    curWSht.Cells(rowToIns, 2).value = curForm
    curWSht.Cells(rowToIns, 3).value = rangeValLbl.Caption
    ActiveWorkbook.Close False
    Kill folForWrite & curForm
    Application.EnableEvents = True
    
End Sub


