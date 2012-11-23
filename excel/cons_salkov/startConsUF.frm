VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} startConsUF 
   Caption         =   "Выберите требуемое действие"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   OleObjectBlob   =   "startConsUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "startConsUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public passForCons As Variant

Private Sub cancelBtn_Click()
    
    Unload startConsUF
End Sub

Private Sub newConsBtn_Click()
    Dim tmpCol As New Collection
    Dim flw As New FileWorker
    Dim strSource As String, xmlName As String, tagName As String, sources As New Collection
    
    startConsUF.Hide

    With dataForConsUF
        'disable this buttons coz this is new consolidation and no way to use this buttons
        .nextBtn.Enabled = False
        .remFileBtn.Enabled = False
        .addFileBtn.Enabled = False 'this button remains disabled until structure file won't be specified
        If startConsUF.usePrevCons.value Then
            .strFileName.ForeColor = 0
            tagName = "source"
            xmlName = startConsUF.Tag & "sources.xml"
            Set sources = flw.xmlCrawler(xmlName, tagName)
            strSource = sources(1)
            strSource = flw.extractNameWithExt(strSource)
            .strFileName.Caption = strSource
            .chooseStrFileBtn.Enabled = False
            .addFileBtn.Enabled = True
            .addFileMem.Caption = "Выберите файлы для консолидации" 'display memorize about picking files for consolidation
            .addFileMem.ForeColor = 255
        Else
            .strFileName.ForeColor = 255 'display tip that str file is not picked
        End If
        
    End With
    
    dataForConsUF.Show
End Sub



Private Sub UserForm_Initialize()
    'check if default folder exists if not disables buttons
    Dim userPath As String, appFolderName As String, backupFol As String
    Dim flw As New FileWorker
    
    
    userPath = Environ$("USERPROFILE")
    appFolderName = userPath & "\cons_report_app_output\" 'folder for output from consolidation app, folder places within default user folder
    
    
    'here password for consolidation can be changed
    passForCons = "123"
    
    
    'creates folder if it isn't exist already and disables appropriate buttons
    If flw.PathExists(appFolderName) Then
        'check if the previous consolidation ends with "ok" status
        If Dir(appFolderName & "prev_cons_ok.xml") <> "" Then
            startConsUF.Tag = appFolderName
            With startConsUF.usePrevCons
                .Enabled = True
                .value = True
            End With

        Else
            'if not "ok" delete all help files
            flw.deleteFilesFromFolder (appFolderName)
            'clean backup_files folder and delete it
            backupFol = appFolderName & "backup_files\"
            If flw.PathExists(backupFol) Then
                flw.deleteFilesFromFolder (backupFol)
                ChDir userPath
                Application.DefaultFilePath = userPath
                RmDir backupFol
            End If
            With startConsUF
                .Tag = appFolderName 'store default folder path in a userform tag
                '.usePrevCons.Enabled = False
                '.editConsBtn.Enabled = False
                .usePrevCons.Enabled = False
                .usePrevCons.value = False
            End With
        End If
    Else
        MkDir appFolderName
        With startConsUF
            .Tag = appFolderName 'store default folder path in a userform tag
            '.usePrevCons.Enabled = False
            '.editConsBtn.Enabled = False
            .usePrevCons.Enabled = False
            .usePrevCons.value = False
        End With
    End If
    'delete status from previous consolidation
    If Dir(appFolderName & "prev_cons_ok.xml") <> "" Then
        Kill appFolderName & "prev_cons_ok.xml"
    End If
End Sub
