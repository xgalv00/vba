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
    Dim tmpCell As Range
    Dim tmpAddr As String, tmpStr As String
    Dim srcAddr As String, destAddr As String
    Dim foundCell As Range
    Dim destSht As Worksheet, srcSht As Worksheet
    Dim i As Integer, clmn As String
    
    
    For Each sht In shtsCopy
    
        Set destSht = wbTo.Sheets(sht)
        Set srcSht = wbFrom.Sheets(sht)
        
        For i = 1 To colsTo.Count
            
            clmn = colsTo(i)
            tmpStr = clmn & "1"
            destSht.Activate
            
            Set tmpCell = Range(tmpStr)
            
                
            Do While tmpCell.Row <= destSht.UsedRange.Rows.Count + destSht.UsedRange.Row
                
                If tmpCell.Interior.Color = 13434879 Then
                    destAddr = tmpCell.Address(False, False)
                    srcAddr = colsFrom(i) & tmpCell.Row
                    'condition for productivity does not copy same values (like null or empty)
                    If Not destSht.Range(destAddr).value = srcSht.Range(srcAddr).value Then
                        destSht.Range(destAddr).value = srcSht.Range(srcAddr).value
                    End If
                End If
                
                Set tmpCell = clw.move_down(tmpCell)
                Set tmpCell = Range(clmn & tmpCell.Row)
            Loop
        
        Next i
        
    Next sht

End Sub

