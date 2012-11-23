Attribute VB_Name = "Module1"
Sub addWBActivate()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim formsFolder As String
    Dim tmpColl As Collection
    Dim wBook As Workbook
    Dim filesToPrcs As Collection
    Dim flw As New FileWorker
    Dim cdw As New CodeWorker
    'Dim fName As String
    Dim fullFName As String
    Dim activateExist As Boolean
    
    formsFolder = "C:\Users\GalkinVa\files_for_transport"
    
    Set tmpColl = flw.getPathsToFilesFromFolder(formsFolder)
    
    If tmpColl Is Nothing Then
        Err.Raise 13, Description:="tmpColl variable doesn't set"
    End If
    
    Set filesToPrcs = tmpColl
    
    For Each fName In filesToPrcs
        
        fullFName = fName
        'rewrite coz fName here equals to fullFName
        fName = flw.extractNameWithExt(fullFName)
        Set wBook = Workbooks.Open(fullFName)
        Set VBProj = wBook.VBProject
        'add here check for reference existence
        
        'check if ThisWorkbook or Ёта нига exist
        If cdw.VBComponentExists("ThisWorkbook", VBProj) Then
            Set VBComp = VBProj.VBComponents("ThisWorkbook")
        ElseIf cdw.VBComponentExists("Ёта нига", VBProj) Then
            Set VBComp = VBProj.VBComponents("Ёта нига")
        Else
            Err.Raise 13, "try to set VBComponent", "components from check doesn't exist in given workbook"
        End If
        
        Set codeMod = VBComp.CodeModule
        
        Set tmpColl = cdw.ListProcedures(VBComp)
        
        'add check for tmpColl is nothing
        
        For Each proc In tmpColl
            If proc = "Workbook_Activate" Then
                activateExist = True
            End If
        Next proc
        
        If Not activateExist Then
            Call cdw.CreateEventProcedure(VBComp)
        Else
            Debug.Print "Workbook_Activate already exist in " & wBook.Name
        End If
        wBook.RunAutoMacros xlAutoClose
        On Error Resume Next
        wBook.Close saveChanges:=True
        If Err.Number <> 0 Then
            Debug.Print "Error occured when try to save " & wBook.Name
        End If
    Next fName
    

End Sub




















