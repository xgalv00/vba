VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub SendStatusMail(addresList, text, subject)
    '
    Dim objOutlook As Object
    Dim objOutlookMsg As Object
    Dim objOutlookRecip As Object
    Dim objOutlookAttach As Object
    '
    Set objOutlook = CreateObject("Outlook.Application")
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
    '
    With objOutlookMsg
        '
        .to = addresList
        .subject = subject
        .Body = text
        .Send
    End With
    '
    Set objOutlookMsg = Nothing
    Set objOutlook = Nothing
    Set objOutlookRecip = Nothing
    Set objOutlookAttach = Nothing
    '
End Sub
Sub ShowSendMessage(alert, addres, subject, text)
    '
    Dim f As UserFormMessage
    Set f = New UserFormMessage
    '
    f.LabelMessage = alert
    f.TextBoxSubject = subject
    f.TextBoxText = text
    f.TextBoxTo = addres
    '
    f.Show
    '
End Sub
