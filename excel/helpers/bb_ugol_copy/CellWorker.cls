VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function move_up(moveFromCell As Range, Optional howFar As Integer) As Range
    If howFar = 0 Then
        howFar = 1
    End If
    
    Cells(moveFromCell.Row - howFar, moveFromCell.Column).Select
    Set move_up = ActiveCell
End Function

Public Function move_down(moveFromCell As Range, Optional howFar As Integer) As Range
    Dim cachedSht As Worksheet

    Set cachedSht = ActiveSheet
    
    moveFromCell.Parent.Activate
    
    If howFar = 0 Then
        'if this check isn't present loops infinitely
        If moveFromCell.MergeCells Then
            howFar = moveFromCell.MergeArea.Rows.Count
        Else
            howFar = 1
        End If
    End If
    
    Set move_down = Cells(moveFromCell.Row + howFar, moveFromCell.Column)
    cachedSht.Activate
    
End Function
Public Function move_right(moveFromCell As Range, Optional howFar As Integer) As Range
    Dim cachedSht As Worksheet
    
    If howFar = 0 Then
        howFar = 1
    End If
    
    Set cachedSht = ActiveSheet
    
    moveFromCell.Parent.Activate
    
    Set move_right = Cells(moveFromCell.Row, moveFromCell.Column + howFar)
    
    cachedSht.Activate
    
End Function
Public Function move_left(moveFromCell As Range, Optional howFar As Integer) As Range
    If howFar = 0 Then
        howFar = 1
    End If
    
    Cells(moveFromCell.Row, moveFromCell.Column - howFar).Select
    Set move_left = ActiveCell
End Function

