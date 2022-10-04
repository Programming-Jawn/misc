VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "neighbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public traversed As Boolean
Private neighbors() As String
 

Property Get getNeighbors() As String()
    getNeighbors = neighbors()
End Property

Sub init(currentCell As String)
    traversed = False
    Dim i As Long
    i = 0
    
    Dim rng As Range
    Set rng = Range(currentCell)
    
    If (rng.Borders(xlEdgeBottom).LineStyle <> XlLineStyle.xlContinuous) Then
        ReDim Preserve neighbors(i)
        neighbors(UBound(neighbors)) = rng.Offset(1, 0).Address
        i = i + 1
    End If
    
    If (rng.Borders(xlEdgeTop).LineStyle <> XlLineStyle.xlContinuous) Then
        ReDim Preserve neighbors(i)
        neighbors(UBound(neighbors)) = rng.Offset(-1, 0).Address
        i = i + 1
    End If
    
    If (rng.Borders(xlEdgeRight).LineStyle <> XlLineStyle.xlContinuous) Then
        ReDim Preserve neighbors(i)
        neighbors(UBound(neighbors)) = rng.Offset(0, 1).Address
        i = i + 1
    End If

    If (rng.Borders(xlEdgeLeft).LineStyle <> XlLineStyle.xlContinuous) Then
        ReDim Preserve neighbors(i)
        neighbors(UBound(neighbors)) = rng.Offset(0, -1).Address
        i = i + 1
    End If
End Sub
