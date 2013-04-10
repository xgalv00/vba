Attribute VB_Name = "ws_change_module"
Dim IE As Object
Dim statusText As String
Public statusChanged As Boolean

Sub wsChangePrep(inCompVal As String, inDsVal As String, inTimeVal As String, inStatus As String)
    'prepares given values for use in ws changing procedure
    
    Dim compValue As String, dsValue As String, timeValue As String, status As Integer
    Dim url As String
    
   
    'encode values for Internet usage
    compValue = "%3A" & dummyEncUrl(inCompVal) & "%3B"
    dsValue = "%3A" & dummyEncUrl(inDsVal) & "%3B"
    timeValue = "%3A" & dummyEncUrl(inTimeVal) & "%3B"
    'map rus status to english
    inStatus = rus_to_eng(inStatus)
    statusText = inStatus
    
    'create url
    url = "http://v-sap-qbi/OSOFT/Landing.aspx?PAGEMODE=WORKSTATUS&appset=DTEK&app=CONSOLIDATION&CVDATA=ACTIVITY%3ABA000%3BCategory%3AAD%3BCOMPANY" & compValue & "CONTRACT%3ACON%5FNONE%3BCREDITQUALITY%3AKK%5FALL%3BCREDITRATING%3AKR%5FALL%3BCURRENCY%3ACUR%5FALL%3BC%5FACCT%3ADB101010%3BC%5FACCT%5FC%3ACCOA%5FALL%3BDATASRC" & dsValue & "DEBTSUBJ%3APZ%5FALL%3BFLOW%3ARF%3BGOODS%3AGD%5FALL%3BGROUPS%3ANON%5FGROUP%3BIFRS7%3AZD%5FALL%3BPARTNER%3AP70000002%3BPARTNER%5FC%3ACPAR%5FALL%3BPERIOD1%5FPAY%3APER1%5FALL%3BSEGMENT%3ASG%5FALL%3BSERIESCB%3ASCB%5FALL%3BSTCKKND%3ASTC%5FALL%3BTERM1%5FCLEAR%3ASR1%5FALL%3BTERM2%5FBEGIN%3ASR2%5FALL%3BTERM3%5FOTHER%3ASR3%5FALL%3BTime" & timeValue & "MEASURES%3AYTD"
    status = 0
    If StrComp("APPROVED", inStatus) = 0 Then
        status = 4
    ElseIf StrComp("REJECTED", inStatus) = 0 Then
        status = 3
    ElseIf StrComp("SUBMITTED", inStatus) = 0 Then
        status = 2
    ElseIf StrComp("STARTED", inStatus) = 0 Then
        status = 1
    End If
    
    'call status changing procedure
    Call IE_Automation(url, status)

End Sub
Function make_dictionary(Optional rus_to_eng As Boolean) As Collection
    'if rus_to_eng is true than russian statuses will be keys.
    Dim helperSht As Worksheet
    Dim clw As New CellWorker
    Dim tmpColl As New Collection
    Dim tmpRng As Range
    
    Set helperSht = ActiveWorkbook.Sheets("Helper")
    helperSht.Select
    If rus_to_eng Then
        'russian statuses are keys
        Set tmpRng = Range("A1")
        Do While tmpRng <> ""
            tmpColl.Add Cells(tmpRng.Row, tmpRng.Column + 1).Value, tmpRng.Value
            Set tmpRng = clw.move_down(tmpRng)
        Loop
    Else
        'english statuses are keys
        Set tmpRng = Range("B1")
        Do While tmpRng <> ""
            tmpColl.Add Cells(tmpRng.Row, tmpRng.Column - 1).Value, tmpRng.Value
            Set tmpRng = clw.move_down(tmpRng)
        Loop
    End If
    
    Debug.Assert tmpColl.Count > 0 And tmpColl.Count < 6 'only five statuses are in consolidation
    
    Set make_dictionary = tmpColl
    
End Function
Private Function convert_rus_to_eng(rusStatus As String) As String
    
    Dim rus_eng_dict As Collection
    
    Set rus_eng_dict = make_dictionary(True)
    
    rus_to_eng = rus_eng_dict(rusStatus)
End Function

Private Function bulkAddToCol(ParamArray Vals() As Variant) As Collection
    'collection must be the first argument
    Dim tmpCol As Collection
    Dim i As Integer
    
    Set tmpCol = Vals(0)
    
    For i = 1 To UBound(Vals)
        tmpCol.Add (Vals(i))
    Next i

    If Not tmpCol Is Nothing Then
        Set bulkAddToCol = tmpCol
    Else
        Err.Raise 9, , "bulkAddToCol: First argument wasn't a collection object or I am not working properly"
    End If
End Function


Private Function dummyEncUrl(valToEncode As String) As String
    'replace . and _ only
    Dim res As Long
    
    res = InStr(valToEncode, ".")
    If res <> 0 Then
        valToEncode = Replace(valToEncode, ".", "%2E")
    End If
    res = 0
    res = InStr(valToEncode, "_")
    If res <> 0 Then
        valToEncode = Replace(valToEncode, "_", "%5F")
    End If
    
    dummyEncUrl = valToEncode
End Function

'private function openUrl(url as String

Private Sub IE_Automation(url As String, status As Integer)
    Dim i As Long
    Dim objElement As Object
    Dim objCollection As Object
    Dim isStatusChanged As Boolean
    
    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
    
    Debug.Assert Not IE Is Nothing 'IE should be installed on user's computer and working properly
    
    ' Send the work status data To URL As GET request
    IE.Navigate url
     
     ' You can uncoment Next line To see work status results
    IE.Visible = True
    
    ' Wait while IE loading...
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
    'clicks one time on a green arrow within first screen
    Call moveToStatusSelection
    
    'waits until select element will be loaded
    Set objCollection = myGetByID("WShselStatus")
    
    'choose appropriate status for submit
    objCollection.selectedIndex = status
    'submit button object
    Set objCollection = myGetByID("imgSp40607")
    objCollection.Disabled = False 'without this Approved status selection doesn't work properly
    objCollection.Click
    
    'sanity check of actual status changing
    statusChanged = isStatusChanged()
 
    ' Clean up
    IE.Quit
    Set IE = Nothing
    Set objCollection = Nothing
 
End Sub
Private Function myGetByID(idName As String) As Object
    '@todo write normal function that will initialize html elements by given id
    Dim htmlElem As Object
    
    '@todo add time checking for more stability
    Do While htmlElem Is Nothing
        Set htmlElem = IE.document.getElementByID(idName)
    Loop
    
    Set myGetByID = htmlElem
End Function

Private Sub moveToStatusSelection()
    'Emulates click on use current cv or another cv screen
    
    Dim objCollection As Object
    
    Set objCollection = myGetByID("imgSp406")
    
    Debug.Assert Not objCollection Is Nothing 'If this object is nothing means that either this object doesn't exist
    'within BPC portal or user have some problems with access to it
    
    objCollection.Click
End Sub

Private Function isStatusChanged() As Boolean
    'Returns true if after status changing procedure name of status changed too
    
    Dim tmpStr As String


    Call moveToStatusSelection
    
    Set objCollection = myGetByID("WShtabCurStatus")
    
    Debug.Assert Not objCollection Is Nothing
    
    tmpStr = objCollection.innerText
    'statusText is global variable and contains inStatus parameter from wsChangePrep.
    If InStr(1, tmpStr, statusText) <> 0 Then
        isStatusChanged = True
    End If
    
End Function

