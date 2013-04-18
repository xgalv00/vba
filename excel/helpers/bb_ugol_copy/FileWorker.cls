VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function getFolderPath(dialogTitle As String) As String
    Dim tmpFolderPath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = dialogTitle
        .Show
        If .SelectedItems.Count <> 0 Then
             tmpFolderPath = .SelectedItems(1)
             getFolderPath = tmpFolderPath
        End If
    End With
End Function

Public Function extractPath(fullFileName As String) As String
    
    'Returns the path from a path\filename string with last \

    Dim temp As Variant
    Dim pathLength As Integer
    Dim fileNameWithExt As String
    
    temp = Split(fullFileName, Application.PathSeparator)
    fileNameWithExt = temp(UBound(temp))
    pathLength = Len(fullFileName) - Len(fileNameWithExt)
    extractPath = Left(fullFileName, pathLength)
    
End Function

Public Function extractName(fullFileName As String) As String

    'Returns file name without path and extension
    
    Dim temp As Variant
    Dim fileNameLength As Integer
    Dim fileNameWithExt As String
    Dim ext As String
    
    temp = Split(fullFileName, ".")
    ext = temp(UBound(temp))
    fileNameLength = Len(fullFileName) - Len(ext) - 1
    temp = Split(Left(fullFileName, fileNameLength), "\")
    extractName = temp(UBound(temp))

End Function

Public Function extractExt(fullFileName As String) As String
    
    'Returns the extension from path\filename string
    
    Dim temp As Variant
        
    temp = Split(fullFileName, ".")
    extractExt = temp(UBound(temp))
        
End Function

Public Function extractNameWithExt(fullFileName As String) As String
    'return name with extention
    extractNameWithExt = extractName(fullFileName) & "." & extractExt(fullFileName)
End Function

Function PathExists(pname) As Boolean
    ' Returns TRUE if the path exists
    On Error Resume Next
    PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function
Private Function resolveFolName(folName As String) As String
    'adds \ to the folder name if it isn't already on place
    Dim tmpStr As String
    tmpStr = Right(folName, Len(folName) - (Len(folName) - 1))
    If tmpStr = "\" Then
        tmpStr = folName
    Else
        tmpStr = folName & "\"
    End If
    
    resolveFolName = tmpStr
End Function

Public Function getPathsToFilesFromFolder(folName As String) As Collection
    'returns collection of pathes to files from given folder
    Dim tmpColl As Collection
    Dim f As String
    
    folName = resolveFolName(folName)
    
    If Not PathExists(folName) Then
        MsgBox "Function getPathsToFilesFromFolder: Folder doesn't exits", vbCritical, "Error"
    End If
    
    f = Dir(folName, vbReadOnly + vbHidden + vbSystem)
    Set tmpColl = New Collection
    Do While f <> ""
        'add file path to collection
        tmpColl.Add folName & f
        ' Get next file
        f = Dir()
    Loop
    
    If tmpColl.Count > 0 Then
        Set getPathsToFilesFromFolder = tmpColl
    End If
    
End Function

Public Sub deleteFilesFromFolder(folName As String)
    'clear specified folder and delete this folder
    Dim f As String
    f = Dir(folName, vbReadOnly + vbHidden + vbSystem)
    Do While f <> ""
        'delete file
        Kill folName & f
        ' Get next file
        f = Dir()
    Loop
    'RmDir folName
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

