Attribute VB_Name = "bbUgol_copyPaste"
Dim srcWB As Workbook, destWB As Workbook
Dim srcWSht As Worksheet, destWSht As Worksheet
Dim relToRange  As String

Sub startPoint()

    '
    Dim ctrlSht As Worksheet
    Dim ctrlRng As Range
    Dim relToRng As Range
    Dim addrForCopy As String, relToAddr As String
    
    Set srcWB = Workbooks("model_in.xlsm")
    Set destWB = Workbooks("model_out.xlsm")
    
    Set srcWSht = srcWB.Sheets("ÁÏÑÑ_ø")
    Set ctrlSht = srcWB.Sheets("control_table_ÁÏÑÑ_ø")
    Set destWSht = destWB.Sheets("ÁÏÑÑ_ø")
    
    Set ctrlRng = ctrlSht.Range("E3")
    Set relToRng = ctrlSht.Range("B3")
    
    relToRange = relToRng.value
    addrForCopy = ctrlRng.value
    
    Call copyRowOfRanges(addrForCopy)
    
End Sub

'Open files for copy

'Find needed mine or it's range

'Compute range address for copying

'Copy one range to another
Private Sub copyRowOfRanges(inRange As String)
    '
    Dim addrForCopy As String
    
    addrForCopy = convertToA1(inRange)
    
    destWSht.Range(addrForCopy).Value2 = srcWSht.Range(addrForCopy).Value2
    
    'copyRowOfRanges(nextRange) - recursive call

End Sub

Private Function convertToA1(inRange As String) As String
    '
    '(str)->str
    
    'Returns converted inRange address to xlA1 style
    
    'relToRange="E149"
    '>>>convertToA1(R[6]C[2]:R[7]C[3])
    '"G155:H156"
    convertToA1 = Application.ConvertFormula(inRange, xlR1C1, xlA1, , relToRange)
    
End Function


''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
Function returnRangeAddr(tmpRange As Range) As String

    returnRangeAddr = tmpRange.Address(False, False)

End Function

Function convertToR1C1(tmpRange As Range, relativeTo As Range) As String

    Dim tmpString As String
    '>>>convertToR1C1(Range("G17:H18"),Range("E9"))
    'R[8]C[2]:R[9]C[3]
    'Debug.Print ""
    tmpString = tmpRange.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, relativeTo:=relativeTo)
    convertToR1C1 = tmpString
End Function
