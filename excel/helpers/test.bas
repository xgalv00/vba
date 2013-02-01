Attribute VB_Name = "test"
Sub test()
    Dim c As Object
    
    'sets vba extesibility library without error messages
    Call setVBAExtensibility
    
    For Each c In ThisWorkbook.VBProject.VBComponents
    
        If InStr(1, c.Name, "CodeWorker") <> 0 Then ThisWorkbook.VBProject.VBComponents.Remove c
    
    Next c
    
    start_time = Now()
    
    ThisWorkbook.VBProject.VBComponents.Import "C:\Users\GalkinVa\Documents\my_macroses\excel\helpers\CodeWorker.cls"
    
    end_time = Now()
    Debug.Print DateDiff("s", start_time, end_time)
    
    Call importTest
    

End Sub

Sub importTest()
    Dim cdw As New CodeWorker
    
   cdw.ComponentTypeToString(

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
