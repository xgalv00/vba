Attribute VB_Name = "Module1"
Sub writeReportName(fName As String, folName As String)
    
     
    UserForm1.Show False
    UserForm1.setBooks fName, folName
    'UserForm1.flagEnd = False
    'UserForm1.setBooks fName, folName
    While UserForm1.rangeValLbl = ""
        Application.Wait DateAdd("s", 5, Now)
    Wend
    Unload UserForm1
    
End Sub
Sub writeReport()
    
    Call writeReportDir("C:\Users\GalkinVa\Desktop\all_forms\œ–Œ¬≈– »\")

End Sub

Sub writeReportDir(folPath As String)
    Dim fName As String
    
    fName = Dir(folPath)
    Do While fName <> ""
        writeReportName fName, folPath
        fName = Dir("")
    Loop
End Sub
