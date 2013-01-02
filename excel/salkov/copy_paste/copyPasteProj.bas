Attribute VB_Name = "copyPasteProj"
Dim colsFrom As Collection, colsTo As Collection, shtsCopy As Collection

Sub stub()
    Dim wbFrom As Workbook, wbTo As Workbook, colsFrom As Collection, colsTo As Collection, shtsCopy As Collection
    Dim fPath As String
    
    fPath = "C:\Users\GalkinVa\Desktop\column_copy_mapping.txt"
    
    Call readMapFromFile(fPath)
    Call copyColumns(wbFrom, wbTo)

End Sub

Sub readMapFromFile(fPath As String)
    Dim flw As New FileWorker
    Dim tmpString As String
    Dim tmpColl As Collection
    
    If Dir(fPath) <> "" Then
        
        Set tmpColl = flw.readLinesFromTxt(fPath) 'call to function that reads file line by line and returns its content like collection of lines
        
    Else
        MsgBox "Check if following path exists " & fPath
    End If
    

End Sub

Sub copyColumns(wbFrom As Workbook, wbTo As Workbook)

    


End Sub

