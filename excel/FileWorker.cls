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
        .title = dialogTitle
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
    extractName = Left(fullFileName, fileNameLength)

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

Public Sub createXml(dFileName As String, rootTagName As String)
    'create first lines of xml document and save this file
    'opens text file
    
    Open dFileName For Output As #1
    
    'Write the <xml> tags
    Print #1, "<?xml version=""1.0"" encoding=""Windows-1251"" standalone=""yes""?>"
    Print #1, "<" & rootTagName & "_list xmlns:xsi=""http://www.w3.org/2001/XMLSchemainstance"">"
    
    'close xml file
    Close #1
End Sub


Public Sub appendLineToXml(dFileName As String, tagName As String, Optional value As String, Optional Level As Integer)
    'create xml line from function argument and append it to the end of the specified file
    'Dim fileSize As Long
    Dim foundLine As String, lineForSearch As String, tagValue As String, lineForFind As String, indent As String
    Dim startTagFound As Boolean, closeTagFound As Boolean
    
    'create indent string based on level of xml record, if level variable is omitted think that value is 1
    'add one for more meaningful way of using this routine
    indent = String(Level + 1, vbTab)
    
    'check whether value for this xml record specified
    'if not print start tag or close tag from given tagName
    If value = "" Then
    
        'decide what type of tag print
        Open dFileName For Input As #1
        closeTagFound = False
        startTagFound = True
        
        lineForFind = indent & "<" & tagName & ">"
        Do Until EOF(1)
            Line Input #1, lineForSearch
            If InStr(1, lineForSearch, lineForFind) Then
                'if you are here that means start tag was found and now you must looking for close tag
                'that is why startTagVariable variable is true from the beginning
                If startTagFound Then
                    'change line that must be found
                    lineForFind = indent & "</" & tagName & ">"
                    startTagFound = False
                    closeTagFound = True
                Else
                    lineForFind = indent & "<" & tagName & ">"
                    startTagFound = True
                    closeTagFound = False
                End If
            End If
            If EOF(1) Then
                If Not closeTagFound Then
                    startTagFound = False 'this condition evaluate when no tag was found
                End If
            End If
            
        Loop
        
        Close #1
        
        Open dFileName For Append As #1
        
        If Not startTagFound And Not closeTagFound Then
            Print #1, indent & "<" & tagName & ">"
        Else
            Print #1, indent & "</" & tagName & ">"
        End If
    
    Else
        Open dFileName For Append As #1
        'print regular tag and value pair
        Print #1, indent & "<" & tagName & ">" & value & "</" & tagName & ">"
    End If
    'close xml file
    Close #1

End Sub

Public Function closeXml(dFileName As String, rootTagName As String)
    'write root close tag to specified xml document
    
    Open dFileName For Append As #1
    
    'Write close tag
    Print #1, "</" & rootTagName & "_list>"
    
    'close xml file
    Close #1


End Function

'@todo think how to make this function as reusable as possible
Public Function extractXMLValue(tagName As String, xmlName As String, Optional rootTagName As String) As Collection
'retrieves value that corresponds to tag name from xml document. If many nodes exist can retrieve value within some root tag name
'maybe replace root tag name by hierarhy level

'@todo find in file tag name from arguments list, take whole line with this tag name, trim tags and return result value
    Dim foundLine As String, lineForSearch As String, tagValue As String
    Dim values As New Collection
    
    Open xmlName For Input As #1
    
    Do Until EOF(1)
        Line Input #1, lineForSearch
        If InStr(1, lineForSearch, tagName) Then
            foundLine = lineForSearch
            'clear start tag
            tagValue = Right(foundLine, Len(foundLine) - InStr(foundLine, ">"))
            'clear end tag
            tagValue = Left(tagValue, InStrRev(tagValue, "<") - 1)
            i = values.Count
            values.Add tagValue
        End If
    Loop
    'returns tag value
    Set extractXMLValue = values
    
    Close
End Function

Public Function xmlCrawler(xmlName As String, rootTagName As String) As Collection
    Dim startTagFound As Boolean, endTagFound As Boolean, tagValues As New Collection, tagValue As String
    Dim skipOneLine As Boolean
    startTagFound = False
    endTagFound = False
    skipOneLine = False
    
    Open xmlName For Input As #1
    
    Do Until EOF(1)
        Line Input #1, lineForSearch
        'find start of xml node block
        If InStr(1, lineForSearch, "<" & rootTagName & ">") Then
            foundLine = lineForSearch
            tagValue = Right(foundLine, Len(foundLine) - InStr(foundLine, ">"))
            'clear end tag
            tagValue = Left(tagValue, InStrRev(tagValue, "<") - 1)
            'adds tag value to values collection
            tagValues.Add tagValue

        End If
        'collect all tag values that laid in between of found xml block
    Loop
    
    Close #1
    
    Set xmlCrawler = tagValues
End Function
