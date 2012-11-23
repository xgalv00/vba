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
    
    If howFar = 0 Then
        howFar = 1
    End If
    
    Cells(moveFromCell.Row + howFar, moveFromCell.Column).Select
    Set move_down = ActiveCell
End Function
Public Function move_right(moveFromCell As Range, Optional howFar As Integer) As Range
    
    If howFar = 0 Then
        howFar = 1
    End If
    Cells(moveFromCell.Row, moveFromCell.Column + howFar).Select
    Set move_right = ActiveCell
    
End Function
Public Function move_left(moveFromCell As Range, Optional howFar As Integer) As Range
    If howFar = 0 Then
        howFar = 1
    End If
    
    Cells(moveFromCell.Row, moveFromCell.Column - howFar).Select
    Set move_left = ActiveCell
End Function

