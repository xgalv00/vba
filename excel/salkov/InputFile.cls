VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'variables that stores properties
Dim applNames As Collection
Public dontAddToList As Boolean
Public fullFileName As String
Public isStructureFile As Boolean 'need for check if it is general input file or structure file


'variables for applyNames function and its helper function determineRange, findUpperLeftCell, resolveName

Dim totalAppliedNames As Integer
'Dim lastUsedRange As Range
'Dim wBook As Workbook
Dim wSheet As Worksheet
Dim clr As Double
Dim usedClmnsCnt As Long, usedRowsCnt As Long 'total numbers of rows and columns in used range
Dim firstRow As Long, firstCol As Long 'first number of rows and columns in used range
Dim workRangeAddr As String 'string for findInnerUpperLeftCell
Dim reachedEndRowOfUsedRange As Boolean
Dim reachedEndColOfUsedRange As Boolean 'varibles that shows where to use another counter in resolveName() routine
Dim reachedEndOfWbook As Boolean 'variable that stops loop in applyNames procedure
Dim horEdgeCellAddr As String, vertEdgeCellAddr As String 'corner cells in range for selection
Dim startAddrHor As String, startAddrVert As String 'cell for start searching within used range
Dim shtNum As Integer 'number for use in collection choosing
Dim rowCounter As Long, colCounter As Long  'helper variables that count how many names was applied and where


Property Let setInputFile(fileName As String)
    'set for all validation procedures, for openning files, for applying names procedure
    '@todo add help xml files creation procedures
    Dim userPath As String
    Dim appFolderName As String
    Dim pathInFile As String
    Dim nameInFile As String
    Dim flw As New FileWorker
    Dim sourcesXml As String
    Dim wBook As Workbook
    Dim xt As String
    
    fullFileName = fileName
    pathInFile = extractPath
    dontAddToList = False
    ChDir pathInFile 'changes path that open by default in getOpenFileName method
    appFolderName = startConsUF.Tag 'take path to default folder
    sourcesXml = startConsUF.Tag & "sources.xml"
    
    xt = extractExt
    If Not xt = "xlsx" Or xt = "xlsm" Or xt = "xltm" Or xt = "xltx" Then
        dontAddToList = True
        MsgBox "Вы добавили недопустимый тип файла. Возможно добавление только файлов Excel 2007,2010"
        Exit Property
    End If
    'pathInFile = extractPath()
    'if input file marked as structure it would be renamed and stored in the root folder if not stored in backup_files folder
    'and not renamed
    If Not isStructureFile Then
        nameInFile = extractName & "." & extractExt
        appFolderName = appFolderName & "backup_files\"
    Else
        nameInFile = "structure_file." & extractExt
        'creates new mapping list if it is not already exist
        flw.createXml sourcesXml, "mapping"
        
        isStructureFile = False
    End If
                            '@todo move this checks to the upper level functions
    
    'copy file in default application folder if it isn't exist already
    If Not Dir(fullFileName) <> "" Then
        Debug.Print fullFileName & " does not exist! Source File Missing"
    ElseIf Not Dir(appFolderName & nameInFile) <> "" Then
        FileCopy fullFileName, appFolderName & nameInFile
        'clean from macros
        Workbooks.Open appFolderName & nameInFile
        Set wBook = Workbooks(nameInFile)
        'temporary turn off warning about overwriting files
        Application.DisplayAlerts = False
        wBook.SaveAs appFolderName & nameInFile, 51
        Application.DisplayAlerts = True
        wBook.Close
        'write appropriate info to sources.xml
        If Dir(sourcesXml) <> "" Then
            flw.appendLineToXml sourcesXml, "mapping"
            flw.appendLineToXml sourcesXml, "source", fullFileName, 1
            flw.appendLineToXml sourcesXml, "destination", appFolderName & nameInFile, 1
            flw.appendLineToXml sourcesXml, "mapping"
        Else
            Debug.Print "sources.xml is missing!"
        End If
    Else
        '@todo handle situation when file names will match but paths will be different
        MsgBox "Вы уже добавляли файл с таким именем. Пожалуйста, переименуйте добавляемый файл"
        dontAddToList = True
    End If
    
    
End Property


Property Get getFullFileName() As String
    getFullFileName = fullFileName
End Property

Property Get getNamesFromFile() As Collection
'returns collection of strings that was applied by applyNames routine
    Dim namesXml As String, tagName As String
    Dim flw As New FileWorker
    
    tagName = "range_name"
    namesXml = startConsUF.Tag & "names.xml"
    'fork that allows to use names from names.xml
    If startConsUF.usePrevCons.value Then
        'if consolidation procedure runs more than one names.xml must exist and procedure takes appropriate names from it
        Set applNames = flw.extractXMLValue(tagName, namesXml)
    End If
    Set getNamesFromFile = applNames

End Property


Public Function extractPath() As String
    
    'Returns the path from a path\filename string with last \

    Dim temp As Variant
    Dim pathLength As Integer
    Dim fileNameWithExt As String
    
    temp = Split(fullFileName, Application.PathSeparator)
    fileNameWithExt = temp(UBound(temp))
    pathLength = Len(fullFileName) - Len(fileNameWithExt)
    extractPath = Left(fullFileName, pathLength)
    
End Function

Public Function extractName() As String

    'Returns file name without path and extension
    
    Dim temp As Variant
    Dim fileNameLength As Integer
    Dim fileNameWithExt As String
    Dim ext As String
    
    temp = Split(fullFileName, Application.PathSeparator)
    fileNameWithExt = temp(UBound(temp))
    temp = Split(fileNameWithExt, ".")
    ext = temp(UBound(temp))
    fileNameLength = Len(fileNameWithExt) - Len(ext) - 1
    extractName = Left(fileNameWithExt, fileNameLength)

End Function

Public Function extractExt() As String
    
    'Returns the extension from path\filename string
    
    Dim temp As Variant
        
    temp = Split(fullFileName, ".")
    extractExt = temp(UBound(temp))
        
End Function

Public Function applyNames(wBookName As String)
'entry point procedure for this class
'@todo write Documenter class and class that will collect "@todo" from code

    Dim cell As Range
    Dim cellAddr As String 'use for storing result of findUpperLeftCell()
    Dim rngAddr As String 'use for storing result of determineRange(cell)
    Dim name As String 'name for applying after execution of resolveName()
    Dim findRange As Range 'variable for storing used range
    Dim fullRngAddr As String
    Dim flw As New FileWorker
    Dim rootTagName As String, dFileName As String, tagName As String
    
                                                    '@todo line about start processing workbook
    
    shtNum = 1
    clr = 13434879
    Set applNames = New Collection
    dFileName = startConsUF.Tag & "names.xml"
    rootTagName = "names"
    tagName = "range_name"
    
    flw.createXml dFileName, rootTagName
    
    Do While Not reachedEndOfWbook
        Set wSheet = Workbooks(wBookName).Sheets(shtNum) 'test line
        wSheet.Activate
        Set findRange = wSheet.usedRange
        workRangeAddr = findRange.Address(False, False)
                                    '@todo add here line that appends to step_controller line about processed sheet
        reachedEndColOfUsedRange = False
        reachedEndRowOfUsedRange = False
        'findRange.Select
        cellAddr = findUpperLeftCell(findRange)  'find first cell that corresponds to specified criteria
        'check if on specified sheet appropriate cell was found if not move to next sheet
        If cellAddr = "" Then
            reachedEndColOfUsedRange = True
            reachedEndRowOfUsedRange = True
            shtNum = shtNum + 1
            If shtNum = ActiveWorkbook.Sheets.Count + 1 Then
                reachedEndOfWbook = True
                Debug.Print "End of workbook " & ActiveWorkbook.name & " reached"
                                    '@todo add here line that no appropriate cells was found for this sheet
                                    'add line that workbook is processed
            End If

                                    '@todo add here line that no appropriate cells was found for this sheet
        Else
            Set cell = Range(cellAddr)
        End If
        
        Do While Not (reachedEndColOfUsedRange And reachedEndRowOfUsedRange)
            
            rngAddr = "=" & determineRange(cell)     'selects range of cells from cell returned by findUpperLeftCell function
            
            'Workbooks("test_bar.xlsx").Sheets("БАР").Range(rngAddr).Select 'test line
            
            name = resolveName()
            Debug.Print name                '@todo replace this print by writing to step_controller and add also range address
            ActiveWorkbook.names.Add name:=name, RefersTo:=rngAddr '@todo replace by wBook variable
            'ActiveWorkbook.Names(name).RefersTo = fullRngAddr
            
            For Each Item In names
                'tmpStr = Item.RefersTo
                If Item.name = name Then
                    totalAppliedNames = totalAppliedNames + 1
                    applNames.Add name
                    flw.appendLineToXml dFileName, tagName, name
                    'Range(Item.RefersTo).Select 'test line
                End If
            Next Item
            
            '@todo add line that will add new name to Names collection of Workbook object
            '@todo add condition checking for new name existence in Names collection if true increment totalAppliedNames
            cellAddr = findInnerUpperLeftCell() 'return address of next cell that corresponds to specified criteria within work range
            'On Error Resume Next
            Set cell = Range(cellAddr)
                        
        Loop
                                        '@todo line about end of processed sheet
        'Exit Do 'test line
        
    Loop
                                                    '@todo line about end of processed workbook
    flw.closeXml dFileName, rootTagName
    
    Debug.Print "Quantity of total applied names is " & totalAppliedNames
End Function

Private Function determineRange(cell As Range) As String
    'Determine range which suites to some conditions
    '@todo refactor this function more specifically whether it needs to check total applyed names quantity
    
    Dim i As Long, j As Long 'temp variables for cycle
    Dim workRange As Range, tmpCell As Range
    Dim cellAddr As String, result As String
    Dim horEdgeFound As Boolean, vertEdgeFound As Boolean
    
    
    cell.Select
'    Set workRange = Range(workRangeAddr)
    horEdgeFound = False
    vertEdgeFound = False
    
    
        If Not horEdgeFound Then
        
            i = 0
            j = 0
            'loop to the right until find end of colored edge
            Do While clr = ActiveCell.Interior.Color
                i = i + 1
                startAddrHor = cell.Offset(j, i).Address(rowabsolute:=False, columnabsolute:=False)
                Range(startAddrHor).Select
            Loop
            'returns address of last cell in a row with specified condition
            horEdgeCellAddr = ActiveCell.Offset(0, -1).Address(rowabsolute:=False, columnabsolute:=False)
            horEdgeFound = True
            Set tmpCell = Range(horEdgeCellAddr)
            If horEdgeFound And Not vertEdgeFound Then
                tmpCell.Select
                'loop down from last cell in a row with specified condition
                'until find end of colored edge
                Do While clr = ActiveCell.Interior.Color
                    i = tmpCell.Column - cell.Column 'determines how many columns to left move from given upper left cell
                    j = j + 1
                    cell.Offset(j, i).Select '@todo add vertical edge cell address retriever
                    
                Loop
                'returns address of last cell in a column with specified condition
                'this allowes us get address for range select
                vertEdgeCellAddr = ActiveCell.Offset(-1, 0).Address(rowabsolute:=False, columnabsolute:=False)
                vertEdgeFound = True
            End If
        End If
        
        
        Set tmpCell = Range(vertEdgeCellAddr)
        result = Range(cell, tmpCell).Address
        determineRange = result
        
        Set tmpCell = Range(result)
        tmpCell.ClearContents
        
        
        'fillRange tmpCell 'test line
        

End Function

Private Function findUpperLeftCell(findRange As Range) As String
'determines upper left cell in a given used range
    
    Dim i As Long, j As Long 'temp variables for cycle
    Dim tmpClr As Double
    Dim usedClmnsCnt As Long, usedRowsCnt As Long 'total numbers of rows and columns in used range
    Dim firstRow As Long, firstCol As Long 'first number of rows and columns in used range

    
        With findRange
            usedClmnsCnt = .Columns.CountLarge
            usedRowsCnt = .Rows.CountLarge
            .Select 'test line
            
        End With
        'Set tmpCellTest = Workbooks("test_bar.xlsx").Sheets("БАР").Range("CX1")
        j = 1
        
        'Debug.Print "color value from findUpperLeftCell " & clr
        
        Do While j < usedRowsCnt + 1 'clarify this conditions
            i = 1
            Do While i < usedClmnsCnt + 1
                
                tmpClr = findRange.Cells(j, i).Interior.Color
                findRange.Cells(j, i).Select 'test line
                
                If tmpClr = clr Then
                    'MsgBox "Was found"
                    findUpperLeftCell = findRange.Cells(j, i).Address(rowabsolute:=False, columnabsolute:=False)
                    Exit Function
                End If
                i = i + 1
            Loop
            j = j + 1
        Loop
    
End Function
Private Function findInnerUpperLeftCell() As String
'function determines upper left cell next to selected range
'@todo refactor code of this function to state which will suite for work with applyNames() function

    Dim i As Long, j As Long 'temp variables for cycle
    Dim cell As Range
    Dim determinedRowNum As Long
    Dim newUpperLeftCell As Range, downRightCell As Range
    Dim result As String
    Dim tmpArray As Variant
    Dim workRange As Range
    Dim colCount As Long, firstCol As Long
    
    
    Set cell = Range(startAddrHor)
    cell.Select
    i = 0
    j = 0
    Do While clr <> ActiveCell.Interior.Color
        'maybe insert here condition
        Set workRange = Range(workRangeAddr)
        With workRange
            colCount = .Columns.CountLarge
            firstCol = .Column
        End With
        
        i = i + 1
        startAddrHor = cell.Offset(j, i).Address(rowabsolute:=False, columnabsolute:=False)
        Range(startAddrHor).Select '@todo replace this by variables
        If ActiveCell.Column = (colCount + firstCol - 1) Then '@todo maybe add subtraction by 1
            reachedEndColOfUsedRange = True 'flags about changing row
            
            Set cell = Range(vertEdgeCellAddr) '@todo replace in production function
            
            determinedRowNum = cell.Row + 1
            tmpArray = Split(workRangeAddr, ":") 'maybe we can replace this computing by using downright cell address
            Set downRightCell = Range(tmpArray(1))
            downRightCell.Select 'test line
            Set newUpperLeftCell = Cells(determinedRowNum, firstCol)
            newUpperLeftCell.Select 'test line
            If newUpperLeftCell.Row >= downRightCell.Row Then
                '@todo add condition for moving through the sheets whithin workbook
                reachedEndRowOfUsedRange = True
                Debug.Print "End of the sheet " & ActiveSheet.name & " reached"
                findInnerUpperLeftCell = ActiveCell.Address(rowabsolute:=False, columnabsolute:=False)
               
                shtNum = shtNum + 1
                rowCounter = 1
                colCounter = 0

                '@todo add check for sheet visibility in workbook
                If Not shtNum = ActiveWorkbook.Sheets.Count + 1 Then
                    ActiveWorkbook.Sheets(shtNum).Activate
                Else
                    reachedEndOfWbook = True
                    Debug.Print "End of workbook " & ActiveWorkbook.name & " reached"
                End If
                'sheetNum = ActiveSheet
                Exit Function
            End If
            Set workRange = Range(newUpperLeftCell, downRightCell)
            workRangeAddr = workRange.Address(False, False)
            result = findUpperLeftCell(workRange)
            
                                                            '@todo add check if here will be more rows or columns
                                                            'without data for consolidation maybe it will throw error
            If result = "" Then
                'handle situation when on sheet will be more rows and this rows don't used in consolidation
                reachedEndColOfUsedRange = True
                reachedEndRowOfUsedRange = True
                result = ActiveCell.Address(False, False)
                findInnerUpperLeftCell = result
                shtNum = shtNum + 1
                '@todo line about end of processing worksheet
                If shtNum = ActiveWorkbook.Sheets.Count + 1 Then
                    reachedEndOfWbook = True
                    '@todo line about end of processing workbook
                End If
            Else
                findInnerUpperLeftCell = result
            End If
            
            Exit Function
            '@todo condition which will check if reached down end of used range
        End If
        
    Loop
    
    'Debug.Print "found in a cell " & Workbooks("test_bar.xlsx").Sheets("БАР").Range(horEdgeCellAddr).Address
    findInnerUpperLeftCell = ActiveCell.Address(rowabsolute:=False, columnabsolute:=False)
End Function

Private Function resolveName() As String
'Creates names based on sheet name, horizontal and vertical position of determined range

    
    Dim rngName As String
    
        
    If totalAppliedNames = 0 Then
        'creates first name
        rowCounter = 1
        colCounter = 1
        
        rngName = "sheet" & shtNum & "_row" & rowCounter & "_column" & colCounter
        resolveName = rngName
        
    Else
        If Not reachedEndColOfUsedRange Then
            'change column counter
            colCounter = colCounter + 1
            rngName = "sheet" & shtNum & "_row" & rowCounter & "_column" & colCounter
            resolveName = rngName
        Else
            If Not reachedEndRowOfUsedRange Then
            'change row counter
                colCounter = 1
                rowCounter = rowCounter + 1
                rngName = "sheet" & shtNum & "_row" & rowCounter & "_column" & colCounter
                resolveName = rngName
                reachedEndColOfUsedRange = False
            End If
        
        End If
        
    End If

End Function

Private Sub cleanRange(rangeForClean As Range)
    'cleans formula from all cells in specified range
    '@todo check if some cell is empty or text
    rangeForClean.Copy
    rangeForClean.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub


Private Sub fillRange(rangeForFill As Range)
    'fill given range by static random data
    rangeForFill.value = Rnd() * 1000
End Sub
