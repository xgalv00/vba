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
        '@todo maybe add comments clearing routines
    Else
        MsgBox "Check if following path exists " & fPath
    End If
    
    If Not tmpColl Is Nothing Then
        Call fillCols(tmpColl)
    Else
        MsgBox "Mapping file is empty. Check it, please. Path " & fPath
    End If

End Sub

Sub fillCols(inCol As Collection)

    'helper function that fills collections by appropriate values
    
    Dim fromCol As String, toCol As String, tmpStr As String

    
    Set colsFrom = New Collection
    Set colsTo = New Collection
    Set shtsCopy = New Collection
    
    For Each Item In inCol
        tmpStr = Item
        toCol = Left(tmpStr, InStr(1, tmpStr, "<") - 1)
        fromCol = Right(tmpStr, Len(tmpStr) - InStr(1, tmpStr, "-"))
        
        colsFrom.Add fromCol
        colsTo.Add toCol
        
    Next Item
    
    shtsCopy.Add ("Á_ïðîä")
    shtsCopy.Add ("ÁÏÑÑ")
    shtsCopy.Add ("Óñëóãè_â_ÁÏÑÑ")
    shtsCopy.Add ("Ïðî÷èå_â_ÁÏÑÑ")
    shtsCopy.Add ("ÁÀÐ")
    shtsCopy.Add ("ÁÐÑ")
    shtsCopy.Add ("ÁïÄÐ_60_90")
    shtsCopy.Add ("ÁïÄÐ_110_160")
    
End Sub


Sub copyColumns(wbFrom As Workbook, wbTo As Workbook)

    


End Sub

