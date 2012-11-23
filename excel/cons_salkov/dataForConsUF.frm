VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dataForConsUF 
   Caption         =   "Выберите данные для консолидации"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   OleObjectBlob   =   "dataForConsUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dataForConsUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addFileBtn_Click()
    
    '@todo create mapping.xml and start add values to it from this routine. first item must be structure_file.xlsx mapping
    
    Dim Filt As String
    Dim FilterIndex As Integer
    Dim Title As String
    Dim fileName As Variant
    Dim msg As String
    Dim i As Integer
    Dim path As String, name As String
    Dim ext As String
    Dim inFile As InputFile
    Dim inputFiles As Collection
    Dim pathToBackFiles As String 'path to backup files
    Dim flw As New FileWorker
    Dim dFileName As String 'variable for sources.xml
    
    pathToBackFiles = startConsUF.Tag & "backup_files\"
    
    dFileName = startConsUF.Tag & "sources.xml"
    ' Set up list of file filters
    Filt = "Excel files (*.xlsx;*.xltx;*.xlsm;*.xltm),*.xlsx;*.xltx;*.xlsm;*.xltm"
    ' Display *.* by default
    FilterIndex = 1
    ' Set the dialog box caption
    Title = "Выберите файлы для консолидации"
    ' Get the file name
    fileName = Application.GetOpenFilename _
    (FileFilter:=Filt, _
    FilterIndex:=FilterIndex, _
    Title:=Title, MultiSelect:=True)
    ' Exit if dialog box canceled
    If Not IsArray(fileName) Then
        MsgBox "No file was selected."
        Exit Sub
    End If
    
    'create folder for backup files if it isn't already exist
    If flw.PathExists(pathToBackFiles) Then
        Debug.Print pathToBackFiles & " already exist"
    Else
        MkDir pathToBackFiles
    End If
    
    
    Set inFile = New InputFile
                                                '@todo add check if not allowed file was picked by user
    ' Display name of the file with extention
    For i = LBound(fileName) To UBound(fileName)
        Dim fullFileName As String
        fullFileName = fileName(i)
        inFile.setInputFile = fullFileName
        If Not inFile.dontAddToList Then
            name = inFile.extractName() & "." & inFile.extractExt()
            addedFilesList.AddItem (name)
        End If
    Next i
    remFileBtn.Enabled = True 'enable possibility to delete file from listbox
    If addedFilesList.ListCount > 1 Then
        'enable posibility to go forward and hides reminder about adding files
        nextBtn.Enabled = True
        addFileMem.Caption = ""
        addFileMem.ForeColor = 0
    End If
    
End Sub

Private Sub backBtn_Click()
    Unload dataForConsUF
    startConsUF.Show
End Sub

Private Sub chooseStrFileBtn_Click()
    Dim Filt As String
    Dim FilterIndex As Integer
    Dim Title As String
    Dim fileName As Variant
    Dim msg As String
    Dim i As Integer
    Dim path As String, name As String
    Dim ext As String
    Dim inFile As InputFile
    Dim inputFiles As Collection
    
                                                    '@todo add check if not allowed file was picked by user

    
    ' Set up list of file filters
    Filt = "Excel files (*.xlsx;*.xltx;*.xlsm;*.xltm),*.xlsx;*.xltx;*.xlsm;*.xltm"
    ' Display *.* by default
    FilterIndex = 1
    ' Set the dialog box caption
    Title = "Выберите файл, который может быть эталоном для отчета"
    ' Get the file name
    fileName = Application.GetOpenFilename _
    (FileFilter:=Filt, _
    FilterIndex:=FilterIndex, _
    Title:=Title, MultiSelect:=False)
    ' Exit if dialog box canceled
    If Not fileName <> False Then
        MsgBox "No file was selected."
        Exit Sub
    End If
    
    Set inFile = New InputFile
    'save full structure file name for passing forward
    strFileName.Tag = fileName
    'flag that this is sturcture file
    inFile.isStructureFile = True
    ' Display name of the file in a label
    inFile.setInputFile = fileName
    name = inFile.extractName()
    strFileName.Caption = name
    
    'gives possibility to add files for consolidation
    addFileBtn.Enabled = True
    strFileName.ForeColor = 0 'set color back to black
    addFileMem.Caption = "Выберите файлы для консолидации" 'display memorize about picking files for consolidation
    addFileMem.ForeColor = 255

    
End Sub

Private Sub confConsBtn_Click()
    dataForConsUF.Hide
    Load consOptionsUF
    consOptionsUF.Show
End Sub


Private Sub finishBtn_Click()
    
    Unload startConsUF
    Unload dataForConsUF
End Sub

Private Sub nextBtn_Click()
    'opens files from default app folder and apply names for it and creates consolidation report
    
    '@todo  create step_controller.xml
    '@todo start write to step_controller.xml and use consolidation step to create tmp report

    Dim userPath As String
    Dim outputRepName As String 'name for created report
    Dim filesFolder As String, defAppFolder As String, tmpRepsFolder As String
    Dim procFiles As Integer, consStepVal As Integer
    Dim dFileName As String 'variable for storing name of step_controller.xml
    Dim sourcesXml As String ' variable for storing of sources.xml
    Dim fName As String
    Dim files As New Collection, filesForTmp As New Collection
    Dim reportItem As report
    Dim flw As New FileWorker
    Dim strFile As New InputFile
    Dim tmpPass As String
    
    'fullFileName = fileName
    StartTime = Timer
    tmpPass = startConsUF.passForCons
    
    
    defAppFolder = startConsUF.Tag 'root folder for this app
    'tmpRepsFolder = defAppFolder & "temp_reports\"  'folder for temporary reports storing
    filesFolder = defAppFolder & "backup_files\"   'folder only for files for consolidation storage
    fName = Dir(filesFolder, vbNormal)
    sourcesXml = defAppFolder & "sources.xml"
    
    
    'set structure file full name
    strFile.fullFileName = defAppFolder & "structure_file.xlsx"
    
    'finish creation of sources.xml file
    flw.closeXml sourcesXml, "mapping"
                                    
    dFileName = defAppFolder & "step_controller.xml"
    
    'create step_controller.xml in a default folder and write first lines of it
    flw.createXml dFileName, "step"
        
    Do While fName <> ""
        files.Add filesFolder & fName
        fName = Dir()
    Loop
    
    outputRepName = "tmp_report"
    
    Set reportItem = New report
    
    'return applied names from this structure file
    reportItem.processStrFile strFile
    
    FileCopy strFile.getFullFileName, defAppFolder & outputRepName
    
    
    Workbooks.Open defAppFolder & outputRepName
    Workbooks.Open defAppFolder & "structure_file.xlsx"
    
    'unprotect tmp_report
    For Each sht In Workbooks(outputRepName).Sheets
        sht.Unprotect tmpPass '@todo replace by variable
    Next sht

    
    reportItem.createFastReport files, outputRepName, dFileName, strFile
            
            
    Unload dataForConsUF 'temp line
    
    EndTime = Timer
    
    Debug.Print Format(EndTime - StartTime, "0.0")
    Unload startConsUF
End Sub

Private Sub remFileBtn_Click()
    Dim fileForDel As String
    
    fileForDel = addedFilesList
    addedFilesList.RemoveItem (addedFilesList.ListIndex)
    Kill startConsUF.Tag & "backup_files\" & fileForDel
    
    If Not addedFilesList.ListCount > 1 Then
        nextBtn.Enabled = False
        addFileMem.Caption = "Выберите файлы для консолидации"
        addFileMem.ForeColor = 255
    End If
End Sub
