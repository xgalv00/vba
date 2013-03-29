Attribute VB_Name = "bbUgol_copyPaste"
Sub test()

    Dim tmpRange As Range
    Dim tmpString As String
    Dim anString As String
    
    Set tmpRange = Range("G17:H18")
    tmpString = tmpRange.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, RelativeTo:=Range("E9"))
    anString = Application.ConvertFormula(tmpString, xlR1C1, xlA1, , Range("E9"))
    copyRange (tmpString)

End Sub

'Open files for copy

'Find needed mine or it's range

'Compute range address for copying

'Copy one range to another
Sub copyRange(inRange As String)
    '
    Dim tmpRange As Range
    Set tmpRange = Range(inRange)
    
    
    Range("J17:K18").Value2 = tmpRange.Value2

End Sub

'test line addition
