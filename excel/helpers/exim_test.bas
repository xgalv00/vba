Attribute VB_Name = "exim_test"
Dim folderPath As String
Dim confFileName As String

Sub test()
    'Starting point for export import macro
    Dim eximMode As Boolean 'true is export, false is import
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    eximMode = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    'sets vba extesibility library without error messages
    Call setVBAExtensibility
    
    'global variables assignment.
    'folderPath is where files are located.
    folderPath = "C:\Users\GalkinVa\Documents\my_macroses\excel\helpers\"
    'confFileName is folderPath with the name of config file
    confFileName = folderPath & "conf.txt"
    
    'switch between tested functionality
    If eximMode Then
        For Each comp In ThisWorkbook.VBProject.VBComponents
            If comp.Name <> "exim_test" And (comp.Type <> vbext_ct_Document) Then
                Call exportTest(comp.Name)
            End If
        Next comp
    Else
        Call importTest
    End If
End Sub

Sub exportTest(VBComp As String)
    '
    '(string)->None
    '
    'Takes folderPath from global space where to save exported vbComp.
    'Also there should be 'conf.txt' file where should be added remark about saving
    
    'Call exportTest(folderPath,vbCompName)
    Dim VBCompObj As Object
    Dim nameForConf As String
    
    'replace this assert statement by exception
    Debug.Assert configExists(folderPath)
    '@todo add exception if vbComp with given name isn't exist in this workbook
    Set VBCompObj = ThisWorkbook.VBProject.VBComponents.Item(VBComp)
    
    'name of module with extension
    nameForConf = exportComp(VBCompObj, folderPath)
    
    'Creates a record in conf.txt if needed
    addSaveNote (nameForConf)
    
    'removes module from project
    ThisWorkbook.VBProject.VBComponents.Remove VBCompObj
    
    
End Sub

Sub importTest()

    Dim tmpCol As Collection
    
    '@todo replace by exception
    Debug.Assert configExists(folderPath)
    '@todo replace by exception
    Debug.Assert isConfConsistent()
    
    Set tmpCol = readConfig
    
    Debug.Assert Not tmpCol Is Nothing
    
    'Imports all files that are listed in conf.txt
    For Each tmpItem In tmpCol
        Call importComp(folderPath & tmpItem)
    Next tmpItem

End Sub
Function getFileName(lineFromConfig As Variant) As String
    
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
    'I don't know how to use correctly arrays in this fucking vba!
    'So if you know how to compute length of array please replace this condition
    Debug.Assert UBound(tmpArray) < 2
    If UBound(tmpArray) > 0 Then
        fromFile = tmpArray(1)
    End If
    
    getFileName = fromFile


End Function

Private Function configExists(folderPath As String) As Boolean

    '(str)->bool
    
    'Returns True if file "conf.txt" exists in given folderPath.
    
    '>>>configExists(validPath)
    'True
    '>>>configExists(invalidPath)
    'False
    '>>>configExists(invalidPath)
    'False
    
    Dim confFileName As String
    
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


Private Sub setVBAExtensibility()
    'sets reference to Microsoft VBA Extesibility library by using GUID
    
    If Not isVbaExtesibilitySet Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=5, Minor:=3
    End If
    
    Debug.Assert isVbaExtesibilitySet
End Sub

Private Function isVbaExtesibilitySet() As Boolean

    'Returns True if reference with GUID={0002E157-0000-0000-C000-000000000046}(VBA extensibility) is set
    Dim tmpStr As String
    
    'loop through all refences and look for this with particular GUID
    For Each ref In ThisWorkbook.VBProject.References
    
        tmpStr = ref.GUID
        If tmpStr = "{0002E157-0000-0000-C000-000000000046}" Then
            isVbaExtesibilitySet = True
            Exit For
        End If
    Next ref

End Function

Private Function readConfig() As Collection

    'Reads conf.txt and returns collection of filenames from it
    Dim tmpColl As Collection
    Dim resColl As Collection
    Dim fromFile As String
    
    Set tmpCol = readLinesFromTxt(confFileName)
    If tmpCol Is Nothing Then
        Err.Raise 31037, , "readConfig: File" & folderPath & "conf.txt is empty"
        Exit Function
    End If
    
    Set resColl = New Collection
    'File cleaning from not meaningful symbols
    For Each tmpItem In tmpCol
        fromFile = getFileName(tmpItem)
        Debug.Assert fromFile <> ""
        resColl.Add (fromFile)
    Next tmpItem
    
    Set readConfig = resColl
    
End Function

Private Function readLinesFromTxt(fPath As String) As Collection
    'returns collection of strings line by line, and returns nothing if file is empty
    Dim tmpColl As New Collection
    Dim tmpString As String
    
    Open fPath For Input As #1 ' Open file for input.
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, tmpString ' Read line into temp string.
        tmpColl.Add tmpString
    Loop
    
    Close #1
   
    If tmpColl.Count > 0 Then
        Set readLinesFromTxt = tmpColl
    Else
        Set readLinesFromTxt = Nothing
    End If

End Function

Private Sub addSaveNote(exportedFName As String)

    '(str)->NoneType
    
    'Adds new entry (with exportedFName in it) to "conf.txt"
    
    '>>>addSaveNote(folderPath,exportedFName)
    '
    'If record about this file is already in conf.txt, do nothing
    If isInConfig(exportedFName) Then
        Exit Sub
    End If
    
    Open confFileName For Append As #1
    
    Print #1, "from_file:" & exportedFName;
    
    Close #1
    
End Sub
Private Function isInConfig(fName As String) As Boolean

    'Returns True if fName is already in conf.txt
     
    Dim tmpColl As Collection
    Dim tmpStr As String
    
    'reads all filenames from config file
    Set tmpColl = readConfig
    If tmpColl Is Nothing Then
        Exit Function
    End If
    
    For Each tmpItem In tmpColl
        If tmpItem = fName Then
            isInConfig = True
        End If
    Next tmpItem
    
End Function

Private Function isConfConsistent() As Boolean

    'Returns True if "folderPath & 'conf.txt'" points to existing files
    
    Dim fNames As Collection
    
    Set fNames = readConfig()
    
    For Each fName In fNames
        If Not (Dir(folderPath & fName, vbNormal + vbHidden + vbSystem) <> vbNullString) Then
            Exit Function
        End If
    Next fName
    
    isConfConsistent = True
End Function
    

Private Sub importComp(filePath As String)
    'import one compenent at a time
    ThisWorkbook.VBProject.VBComponents.Import filePath
End Sub

Private Function exportComp(VBComp As Object, _
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

Private Function getCompExtension(VBComp As Object) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            getCompExtension = ".cls"
        Case vbext_ct_Document
            getCompExtension = ".cls"
        Case vbext_ct_MSForm
            getCompExtension = ".frm"
        Case vbext_ct_StdModule
            getCompExtension = ".bas"
        Case Else
            getCompExtension = ".bas"
    End Select
    
End Function
