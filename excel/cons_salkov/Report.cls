VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function processStrFile(strFile As InputFile) As Collection
    'open structure_file.xlsx in default application folder, apply names, clean ranges to which names was applied
    'return collection of applied names maybe with addresses
    Dim strFileFullName As String, strFileName As String
    Dim applNames As New Collection
    Dim namesXml As String, tagName As String, tmpPass As Variant
    Dim flw As New FileWorker
    Dim wSheet As Worksheet
    
    tmpPass = startConsUF.passForCons
    
    
    tagName = "range_name"
    namesXml = startConsUF.Tag & "names.xml"
    strFileFullName = strFile.getFullFileName
    'Workbooks("structure_file.xlsx").Close
    strFileName = strFile.extractName & "." & strFile.extractExt
    
    Workbooks.Open strFileFullName
    Workbooks("structure_file.xlsx").Unprotect tmpPass
    For Each sht In Workbooks("structure_file.xlsx").Sheets
        sht.Unprotect tmpPass
    Next sht
                            '@todo add check if workbook is open
                            
    'fork that allows to use names from names.xml
    If startConsUF.usePrevCons.value Then
        'if consolidation procedure runs more than one names.xml must exist and procedure takes appropriate names from it
        Set applNames = flw.extractXMLValue(tagName, namesXml)
    Else
        'apply names to structure file ranges if names.xml doesn't exist
            strFile.applyNames strFileName
            Set applNames = strFile.getNamesFromFile
    End If
    
    Application.DisplayAlerts = False
    Workbooks("structure_file.xlsx").Protect tmpPass, True
    
    For Each sht In Workbooks("structure_file.xlsx").Sheets
        sht.Protect tmpPass
    Next sht
    
    Workbooks("structure_file.xlsx").SaveAs strFile.getFullFileName, 51
    Workbooks("structure_file.xlsx").Close
    Application.DisplayAlerts = True
    Set processStrFile = applNames
End Function

Public Sub createFastReport(fileNames As Collection, outputRepName As String, dFileName As String, strFile As InputFile, Optional tmpRepFolder As String, Optional outStrFile As InputFile)
    Dim flw As New FileWorker
    'Dim dFileName As String 'variable that represents path to step_controller.xml
    Dim file As InputFile
    Dim tmpName As String, filePath As String
    Dim fileNamesCol As New Collection
    'Dim secondFileNames As Collection
    Dim rngAddr As String
    'Dim firstName As String, secondName As String
    Dim someName As name
    'Dim firstSource As String, secondSource As String, nameRng As String
    Dim tmpArray As Variant 'temp line
    Dim sources As New Collection 'temp line
    Dim arrString As String 'temp line
    Dim rngNames As New Collection
    Dim nameAddr As String, endOfSource As String
    Dim tmpPass As Variant
    'Dim fileName As String
    
    
    Application.Calculation = xlCalculationManual
    tmpPass = startConsUF.passForCons
    flw.appendLineToXml dFileName, "step:validation"
    
    StartTime = Timer
    
    '"'C:\Users\GalkinVa\cons_report_app_output\backup_files\[Модель Бюджетирования для тестов - копия.xlsx]Б_продаж'!R16C7:R27C8"
    For Each fileName In fileNames
        Set file = New InputFile
        file.setInputFile = fileName
        tmpName = file.extractName & "." & file.extractExt
        filePath = file.extractPath
        'apply names to inputted file
                
        Workbooks.Open fileName
    
                                            '@todo add open file line
        Workbooks(tmpName).Activate
        'crete sources strings for consolidation
        sources.Add "'" & filePath & "[" & tmpName & "]"
    Next fileName
    
    Set rngNames = strFile.getNamesFromFile
    
    For Each rngName In rngNames
        Set someName = Workbooks("structure_file.xlsx").names(rngName)
        nameAddr = someName.RefersToR1C1
        nameAddr = Right(nameAddr, Len(nameAddr) - 1)
        tmpArray = Split(nameAddr, "!")
        endOfSource = tmpArray(0) & "'!"
        arrString = ""
        For Each Source In sources
            arrString = arrString & Source & endOfSource & tmpArray(1) & ","
        Next Source
        arrString = Left(arrString, Len(arrString) - 1)
        tmpArray = Split(arrString, ",")
        
        Workbooks(outputRepName).Activate
        Application.Goto reference:=someName.RefersToR1C1
        Selection.Consolidate sources:=tmpArray, Function:=xlSum, TopRow:=False, LeftColumn:=False, CreateLinks:=False

    Next rngName
    '@todo add here code for selection how much files must be consolidated
    '@todo rewrite this for using names from structure file

    Application.DisplayAlerts = False '@todo enable this line
    Workbooks("structure_file.xlsx").SaveAs strFile.getFullFileName, 51
    Workbooks("structure_file.xlsx").Close
    For Each fileName In fileNames
        tmpName = fileName
        tmpName = flw.extractNameWithExt(tmpName)
        Workbooks(tmpName).Close
        Kill fileName
    Next fileName
    Application.DisplayAlerts = True
    
    
    'if this is not set you cannot delete backup_files folder
    Application.DefaultFilePath = Environ("USERPROFILE")
    ChDir Environ("USERPROFILE")
    'RmDir flw.extractPath(strFile.getFullFileName)
    
    'flag about ok status of consolidation
    dFileName = startConsUF.Tag & "prev_cons_ok.xml"
    flw.createXml dFileName, "test"
    flw.closeXml dFileName, "test"
    
    Workbooks(outputRepName).Activate
    Workbooks(outputRepName).Protect tmpPass, True
    For Each sht In Workbooks(outputRepName).Sheets
        sht.Protect tmpPass
    Next sht
    
    Application.Calculation = xlCalculationAutomatic
    EndTime = Timer
    
    Debug.Print "Total time is " & Format(EndTime - StartTime, "0.0")


End Sub

