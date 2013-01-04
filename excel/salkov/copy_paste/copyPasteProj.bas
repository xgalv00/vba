Attribute VB_Name = "copyPasteProj"
Dim colsFrom As Collection, colsTo As Collection, shtsCopy As Collection

Sub startNewCopyBtn_Click()

    colCopyUF.Show

End Sub

Sub stub()
    Dim wbFrom As Workbook, wbTo As Workbook, colsFrom As Collection, colsTo As Collection, shtsCopy As Collection
    Dim fPath As String
    
    fPath = "C:\Users\GalkinVa\Desktop\column_copy_mapping.txt"
    
    Set wbFrom = Workbooks("Модель Бюджетирования_3_для Салькова.xlsx")
    Set wbTo = Workbooks("Модель Бюджетирования_3_для Салькова_копия.xlsx")
    
    Call centralExecUnit(1, wbFrom, wbTo)
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

Public Sub centralExecUnit(checkedVal As Integer, wbFrom As Workbook, wbTo As Workbook)
    Dim inCol As Collection

    'checked values must be in range from 1 to 5 accordingly to options, 1 is Перенести начало периода
    Set inCol = New Collection
    
    Select Case checkedVal
    
        Case 1
            Set inCol = bulkAddToCol(inCol, "J<-DS", "K<-DT", "M<-DV", "N<-DW", "T<-CE", "U<-CF", "W<-CN", "X<-CO", "Z<-CQ", "AA<-CR")
        Case 2
            Set inCol = bulkAddToCol(inCol, "AK<-O", "AL<-P", "BH<-R", "BI<-S", "CE<-U", "CF<-V")
        Case 3
            Set inCol = bulkAddToCol(inCol, "AK<-AA", "AL<-AB", "BH<-AD", "BI<-AE", "CE<-AG", "CF<-AH")
        Case 4
            Set inCol = bulkAddToCol(inCol, "AK<-AM", "AL<-AN", "BH<-AP", "BI<-AQ", "CE<-AS", "CF<-AT")
        Case 5
            Set inCol = bulkAddToCol(inCol, "AK<-AY", "AL<-AZ", "BH<-BB", "BI<-BC", "CE<-BE", "CF<-BF")
    End Select
    
    If inCol.Count = 0 Then
        Set inCol = Nothing
    End If
    
    If Not inCol Is Nothing Then
        Call fillCols(inCol)
    Else
        MsgBox "Что-то не так с кодом, пожалуйста, обратитесь к разработчику"
        Exit Sub
    End If
    
    Call copyColumns(wbFrom, wbTo)
    
End Sub



Public Sub fillCols(inCol As Collection)

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
    
    Set shtsCopy = bulkAddToCol(shtsCopy, "Б_продаж", "БПСС", "Услуги_в_БПСС", "Прочие_в_БПСС", "БАР", "БРС", "БпДР_60_90", "БпДР_110_160")
    
    'shtsCopy.Add ("Б_продаж")
    'shtsCopy.Add ("БПСС")
    'shtsCopy.Add ("Услуги_в_БПСС")
    'shtsCopy.Add ("Прочие_в_БПСС")
    'shtsCopy.Add ("БАР")
    'shtsCopy.Add ("БРС")
    'shtsCopy.Add ("БпДР_60_90")
    'shtsCopy.Add ("БпДР_110_160")
    
End Sub


Public Sub copyColumns(wbFrom As Workbook, wbTo As Workbook)

    Dim clw As New CellWorker
    Dim tmpCell As Range
    Dim tmpAddr As String, tmpStr As String
    Dim srcAddr As String, destAddr As String
    Dim foundCell As Range
    Dim destSht As Worksheet, srcSht As Worksheet
    Dim i As Integer, clmn As String
    
    
    For Each sht In shtsCopy
    
        
        If shtExist(sht, wbTo) Then 'check for skipping optional sheets
        
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
            
        End If
        
    Next sht

End Sub

Private Function bulkAddToCol(ParamArray Vals() As Variant) As Collection
    'collection must be the first argument
    Dim tmpCol As Collection
    Dim i As Integer
    
    Set tmpCol = Vals(0)
    
    For i = 1 To UBound(Vals)
        tmpCol.Add (Vals(i))
    Next i

    If Not tmpCol Is Nothing Then
        Set bulkAddToCol = tmpCol
    Else
        Err.Raise 9, , "bulkAddToCol: First argument wasn't a collection object or I am not working properly"
    End If
End Function

Private Function shtExist(ByVal shtName As String, wb As Workbook) As Boolean
    'returns false if sheet isn't in a given workbook
    
    Dim tmpWSht As Worksheet
    
    On Error Resume Next
    Set tmpWSht = wb.Sheets(shtName)
    On Error GoTo 0
    
    If Not tmpWSht Is Nothing Then
        shtExist = True
    End If
    
    
    
    
End Function

