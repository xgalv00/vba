Attribute VB_Name = "copyPasteProj"
Dim colsFrom As Collection, colsTo As Collection, shtsCopy As Collection

Sub stub()
    Dim wbFrom As Workbook, wbTo As Workbook, colsFrom As Collection, colsTo As Collection, shtsCopy As Collection
    Dim fPath As String
    
    fPath = "C:\Users\GalkinVa\Desktop\column_copy_mapping.txt"
    
    Set wbFrom = Workbooks("Модель Бюджетирования_3_для Салькова.xlsx")
    Set wbTo = Workbooks("Модель Бюджетирования_3_для Салькова_копия.xlsx")
    
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
    
    shtsCopy.Add ("Б_продаж")
    shtsCopy.Add ("БПСС")
    shtsCopy.Add ("Услуги_в_БПСС")
    shtsCopy.Add ("Прочие_в_БПСС")
    shtsCopy.Add ("БАР")
    shtsCopy.Add ("БРС")
    shtsCopy.Add ("БпДР_60_90")
    shtsCopy.Add ("БпДР_110_160")
    
End Sub


Sub copyColumns(wbFrom As Workbook, wbTo As Workbook)

    Dim clw As New CellWorker
    Dim foundCell As Range
    Dim tmpAddr As String
    Dim destSht As Worksheet, srcSht As Worksheet
    Dim tmpRow As Integer
    Dim tmpStr As String
    
    
    For Each sht In shtsCopy
    
        Set destSht = wbTo.Sheets(sht)
        Set srcSht = wbFrom.Sheets(sht)
        
        For Each clmn In colsTo
            tmpStr = clmn & "1"
            'wbTo.Activate
            destSht.Activate
            Range(tmpStr).Select
            
            Application.FindFormat.Interior.Color = 13434879 'here color value can be changed
            
            'move line by line within given column
            Set foundCell = destSht.UsedRange.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=True)
                
            foundCell.Select 'test line
                
            Do While Not foundCell Is Nothing
            
                foundCell.Select 'test line
                
                tmpAddr = foundCell.Address
                'copy only cells that have values
                destSht.Range(tmpAddr).value = srcSht.Range(tmpAddr).value
            
                Set foundCell = ActiveCell.FindNext
            Loop
        
        Next clmn
    Next sht

End Sub

