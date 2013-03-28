Attribute VB_Name = "exim_test"
Sub test()
    Dim c As Object
    Dim folderPath As String
    
    'sets vba extesibility library without error messages
    Call setVBAExtensibility
    
    folderPath = "C:\Users\GalkinVa\Documents\my_macroses\excel\helpers\"
    
    'Call importTest(folderPath)
    Call exportTest(folderPath, "bbUgol_copyPaste")

End Sub

Sub exportTest(folderPath As String, VBComp As String)
    '
    '(string,string)->None
    '
    'Takes folderPath where to save exported vbComp.
    'Also there should be 'conf.txt' file where should be added remark about saving
    
    'Call exportTest(folderPath,vbCompName)
    Dim VBCompObj As Object
    Dim nameForConf As String
    
    Debug.Assert configExists(folderPath)
    
    Set VBCompObj = ThisWorkbook.VBProject.VBComponents.Item(VBComp)
    
    nameForConf = exportComp(VBCompObj, folderPath)
    
    
    
End Sub

Sub importTest(folderPath As String)
    Dim tmpCol As Collection
    Dim confFileName As String
    Dim toComponent As String
    Dim fromFile As String
    'Dim tmpItem As String
    Dim tmpArr As Variant
    Dim tmpArray As Variant
    
    Debug.Assert configExists(folderPath)
    
    confFileName = folderPath & "conf.txt"
    
    Set tmpCol = readLinesFromTxt(confFileName)
    If tmpCol Is Nothing Then
        Err.Raise 31037, , "importTest: File" & folderPath & "conf.txt is empty"
        Exit Sub
    End If
    
    'parse json, maybe move this code to json deserealizing function
    'File cleaning from not meaningful symbols
    For Each tmpItem In tmpCol
        
        fromFile = getFileName(tmpItem)
        Debug.Assert fromFile <> ""
        Call importComp(folderPath & fromFile)
        
    Next tmpItem
    
    

End Sub
Function getFileName(lineFromConfig As String) As String
    
    '(str)->str
    
    'Returns file name from lineFromConfig
    
    '>>>getFileName("from_file:bbUgol_copyPaste.bas")
    'bbUgol_copyPaste.bas
    '>>>getFileName("some text")
    '""
    Dim tmpItem As String, fromFile As String
    Dim tmpArray As Variant
    
    tmpItem = Trim(lineFromConfig)
    tmpArray = Split(tmpItem, ":")
    If UBound(tmpArray) > 0 Then
        fromFile = tmpArray(1)
    End If
    
    lineFromConfig = fromFile


End Function

Function configExists(folderPath As String) As Boolean

    '(str)->bool
    
    'Returns True if file "conf.txt" exists in given folderPath.
    
    '>>>configExists(validPath)
    'True
    '>>>configExists(invalidPath)
    'False
    '>>>configExists(invalidPath)
    'False
    
    configExists = False
    
    confFileName = folderPath & "conf.txt"
    
    If Dir(folderPath) = "" Then 'check for folder existence
        Exit Function
    End If
    
    If Dir(confFileName) = "" Then 'check for config file existence
        Exit Function
    End If

    configExists = True
    
End Function


Sub setVBAExtensibility()
    'sets reference to Microsoft VBA Extesibility library by using GUID
    
    Dim tmpStr As String
    Dim vbaExtesibilitySet As Boolean

    For Each ref In ThisWorkbook.VBProject.References
    
        tmpStr = ref.GUID
        If tmpStr = "{0002E157-0000-0000-C000-000000000046}" Then
            vbaExtesibilitySet = True
            Exit For
        End If
    Next ref
    If Not vbaExtesibilitySet Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=5, Minor:=3
    End If

End Sub

Public Function readLinesFromTxt(fPath As String) As Collection
    'returns collection of strings line by line, and returns nothing if file is empty
    Dim tmpColl As New Collection
    Dim tmpString As String
    
    Open fPath For Input As #1 ' Open file for input.
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, tmpString ' Read line into temp string.
        tmpColl.Add tmpString
        'Debug.Print tmpString ' Print data to the Immediate window.
    Loop
    
    Close #1
   
    If tmpColl.Count > 0 Then
        Set readLinesFromTxt = tmpColl
    Else
        Set readLinesFromTxt = Nothing
    End If

End Function
Sub addSaveNote(folderPath As String, strexportedFName As String)

    '(str)->NoneType
    
    'Adds new entry (with exportedFName in it) to "conf.txt"
    
    '>>>addSaveNote(folderPath,exportedFName)
    '
    
    
End Sub
    

Sub importComp(filePath As String)
    'import one compenent at a time
    ThisWorkbook.VBProject.VBComponents.Import filePath
End Sub

Private Function exportComp(VBComp As VBIDE.vbComponent, _
            FolderName As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim fullFName As String, fName As String
    
    'Extension is computed accrodingly to vbcomponent type
    Extension = getCompExtension(VBComp:=VBComp)
    'use vbcomp's name for saving
    Debug.Assert VBComp.Name <> ""
    fName = VBComp.Name & Extension
    fullFName = FolderName & fName
    
    'Check if file with such a name already exists delete it
    If Dir(fullFName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        Kill fullFName
    End If
    
    VBComp.Export FileName:=fullFName
    exportComp = fName

End Function

Public Function getCompExtension(VBComp As VBIDE.vbComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function
