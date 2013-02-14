Attribute VB_Name = "test"
Sub test()
    Dim c As Object
    Dim folderPath As String
    
    'sets vba extesibility library without error messages
    Call setVBAExtensibility
    
    folderPath = "C:\Users\GalkinVa\Documents\my_macroses\excel\helpers\"
    
    Call importTest(folderPath)
    

End Sub

Sub importTest(folderPath As String)
    Dim tmpCol As Collection
    Dim confFileName As String
    Dim toComponent As String
    Dim fromFile As String
    'Dim tmpItem As String
    Dim tmpArr As Variant
    Dim tmpArray As Variant
    
    confFileName = folderPath & "conf.txt"
    
    If Dir(folderPath) = "" Then 'check for folder existence
        Err.Raise 76, , "importTest: Folder " & folderPath & " doesn't exist"
        Exit Sub
    ElseIf Dir(confFileName) = "" Then 'check for config file existence
        Err.Raise 53, , "importTest: File conf.txt doesn't exist in " & folderPath
        Exit Sub
    End If
    
    Set tmpCol = readLinesFromTxt(confFileName)
    If tmpCol Is Nothing Then
        Err.Raise 31037, , "importTest: File" & folderPath & "conf.txt is empty"
        Exit Sub
    End If
    
    'parse json, maybe move this code to json deserealizing function
    For Each tmpItem In tmpCol
        
        tmpItem = Trim(tmpItem)
        tmpItem = Left(tmpItem, Len(tmpItem) - 1)
        tmpItem = Right(tmpItem, Len(tmpItem) - 1)
        tmpArr = Split(tmpItem, ",")
        tmpArray = Split(tmpArr(0), ":")
        fromFile = tmpArray(1)
        tmpArray = Split(tmpArr(1), ":")
        inComponent = tmpArray(1)
        Call importComp(folderPath & fromFile)
        
    Next tmpItem
    
    

End Sub



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

Sub importComp(filePath As String)
    'import one compenent at a time
    ThisWorkbook.VBProject.VBComponents.Import filePath
End Sub
