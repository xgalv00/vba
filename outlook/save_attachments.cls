VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "save_attachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SaveAttachments()

    Dim objOL As Outlook.Application
    Dim objMsg As Outlook.MailItem 'Object
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long
    Dim msgCount As Integer
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderpath As String
    Dim strDeletedFiles As String
    Dim flw As New FileWorker
    Dim attExt As String, attName As String
    
    
    ' Get the path to your My Documents folder
    
    strFolderpath = Environ("USERPROFILE") & "\" & "OLAttachments\"
    On Error Resume Next

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

    ' Check each selected item for attachments. If attachments exist,
    ' save them to the Temp folder and strip them from the item.
    For Each objMsg In objSelection
        If objMsg.Categories = "Желтая категория" Then
            msgCount = msgCount + 1
            ' This code only strips attachments from mail items.
            ' If objMsg.class=olMail Then
            ' Get the Attachments collection of the item.
            Set objAttachments = objMsg.Attachments
            lngCount = objAttachments.Count
            
            'Use the MsgBox command to troubleshoot. Remove it from the final code.
            'MsgBox objAttachments.Count
            
            If lngCount > 0 Then
            
                 ' We need to use a count down loop for removing items
                 ' from a collection. Otherwise, the loop counter gets
                 ' confused and only every other item is removed.
                 
                 For i = lngCount To 1 Step -1
                 
                     ' Save attachment before deleting from item.
                     ' Get the file name.
                     attName = flw.extractName(objAttachments.Item(i).FileName) & "_" & objMsg.SenderName & msgCount
                     attExt = flw.extractExt(objAttachments.Item(i).FileName)
                     strFile = attName & "." & attExt
                     If attExt = "xlsx" Or attExt = "xlsm" Or attExt = "xls" Or attExt = "xltx" Or attExt = "xltm" Then
                        
                        ' Combine with the path to the Temp folder.
                        strFile = strFolderpath & strFile
                        
                        ' Save the attachment as a file.
                        objAttachments.Item(i).SaveAsFile strFile
                        
                    Else
                        Debug.Print "Attachment " & i & " have extention " & attExt & " from sender " & objMsg.SenderName
                    End If
                 Next i
        
            End If
        End If
    Next
    
ExitSub:
    
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
    Debug.Print "Total processed messages was " & msgCount
End Sub


