Attribute VB_Name = "algorithms"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Dim allCells As Collection
Dim traversed As Collection
Dim startCell As String

'This is the start of the code. Triggered by the form
Sub Driver()
'1.0 setup------------------------------------------------------------

'1.1 initialize
    Dim cell, mz As Range
    
'1.2 set range and collection
    Set mz = Range(Range("B2").Address & ":" & ActiveCell.Address)
    Set allCells = New Collection
    
'1.3 set new class objects inside allCells collections with current cell range as key
    For Each cell In mz
        If (cell.Interior.ColorIndex > 1) Then
            allCells.Add Item:=New neighbor, Key:=cell.Address
            allCells.Item(cell.Address).init (cell.Address)
            If cell.Value = "s" Then
                startCell = cell.Address
            End If
        End If
    Next cell
'2.0 call algorithms-------------------------------------------------

'2.1 BFS
Call BFS

End Sub

Sub enable()
Application.EnableEvents = True

End Sub
Sub BFS()

'enqueue - adds an element at end
'dequeue - removes an element at end
'peek - looks at beginning of queue

    Application.ScreenUpdating = False

    Dim output, queue As Object
    Set output = CreateObject("System.Collections.ArrayList")
    Set queue = CreateObject("System.Collections.Queue")
    Dim v As String
    Dim i As Integer
    
    queue.enqueue startCell
    
    While queue.Count > 0
        v = queue.dequeue
        If Not allCells.Item(v).traversed Then
            output.Add (v)
            allCells.Item(v).traversed = True
            For i = 0 To UBound(allCells.Item(v).getNeighbors())
                If Not allCells.Item(allCells.Item(v).getNeighbors()(i)).traversed Then
                    Dim s As String
                    s = allCells.Item(v).getNeighbors()(i)
                    queue.enqueue s
                End If
            Next i
        End If
    Wend
    
    For i = 0 To output.Count - 1
        Range(output(i)).Select
        ActiveCell.Interior.Color = RGB(HSLtoRGB(Int(i / 8) + 1, 255, 100)(0), HSLtoRGB(Int(i / 8) + 1, 255, 100)(1), HSLtoRGB(Int(i / 8) + 1, 255, 100)(2))
    Next i
    Application.ScreenUpdating = True

End Sub

Sub generate()
    Application.ScreenUpdating = False
    
    Cells.ClearFormats
    Dim cell, mazeRng, lastCol, lastRow As Range
    Dim choose As Long

    Set mazeRng = Range(Range("B2").Address & ":" & ActiveCell.Address)

    Set lastCol = Range(Cells(2, ActiveCell.Column).Address & ":" & ActiveCell.Address)
    Set lastRow = Range(Cells(ActiveCell.Row, 2).Address & ":" & ActiveCell.Address)
    
    For Each cell In mazeRng
        cell.Borders(xlEdgeBottom).LineStyle = xlNone
        cell.Borders(xlEdgeRight).LineStyle = xlNone
    Next cell

    For Each cell In mazeRng
        choose = Rnd()
        If choose = 0 Then
            cell.Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            cell.Borders(xlEdgeBottom).Weight = xlThick
        Else
            cell.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            cell.Borders(xlEdgeRight).Weight = xlThick
        End If
    Next cell
    
    For Each cell In lastCol
        cell.Borders(xlEdgeBottom).LineStyle = xlNone
    Next cell
    
    For Each cell In lastRow
        cell.Borders(xlEdgeRight).LineStyle = xlNone
    Next cell
    
    mazeRng.BorderAround LineStyle:=xlContinuous, Weight:=xlThick

    Application.ScreenUpdating = True

End Sub

Sub generate_SW()
Application.ScreenUpdating = False

Cells.ClearFormats

Dim cell, mazeRng, lastCol, topRow As Range
Dim choose As Long
Set mazeRng = Range(Range("B2").Address & ":" & ActiveCell.Address)
Set lastCol = Range(Cells(2, ActiveCell.Column).Address & ":" & ActiveCell.Address)
Set topRow = Range(Cells(2, 2).Address & ":" & Cells(2, ActiveCell.Column).Address)


'populate all borders
mazeRng.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
For Each cell In mazeRng
    cell.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
Next cell

For Each cell In mazeRng
    choose = Rnd()
    If choose = 0 And Intersect(cell, topRow) Is Nothing Then
        'nothing
    ElseIf Intersect(cell, lastCol) Is Nothing Then
            cell.Borders(xlEdgeRight).LineStyle = xlNone
    End If
Next cell

Dim setOfRun As Object
Set setOfRun = CreateObject("System.Collections.ArrayList")

For Each cell In mazeRng
    
    If (Intersect(topRow, cell) Is Nothing) Then
        setOfRun.Add cell
        If cell.Borders(xlEdgeRight).LineStyle <> xlNone Then
            setOfRun(Int((setOfRun.Count * Rnd) + 0)).Borders(xlEdgeTop).LineStyle = xlNone
            setOfRun.Clear
        End If
        
    End If
    
Next cell
mazeRng.Interior.Color = RGB(80, 150, 50)
Application.ScreenUpdating = True

End Sub

'Sub format(factor)
'MsgBox (factor)
'Dim i As Long
'For i = 1 To 75
'    Columns(i).ColumnWidth = Int((5 - 3 + 1) * Rnd + 3)
'    Rows(i).RowHeight = Int((5 - 3 + 1) * Rnd + 3) * 6.25
'Next i
'End Sub

Function HSLtoRGB(Hue As Integer, Saturation As Integer, Luminance As Integer) As Variant
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    Dim c As Double
    Dim X As Double
    Dim m As Double
    Dim rfrac As Double
    Dim gfrac As Double
    Dim bfrac As Double
    Dim hangle As Double
    Dim hfrac As Double
    Dim sfrac As Double
    Dim lfrac As Double

    If (Saturation = 0) Then
        r = 255
        g = 255
        b = 255
    Else
        lfrac = Luminance / 255
        hangle = Hue / 255 * 360
        sfrac = Saturation / 255
        c = (1 - Abs(2 * lfrac - 1)) * sfrac
        hfrac = hangle / 60
        hfrac = hfrac - Int(hfrac / 2) * 2 'fmod calc
        X = (1 - Abs(hfrac - 1)) * c
        m = lfrac - c / 2
        Select Case hangle
            Case Is < 60
                rfrac = c
                gfrac = X
                bfrac = 0
            Case Is < 120
                rfrac = X
                gfrac = c
                bfrac = 0
            Case Is < 180
                rfrac = 0
                gfrac = c
                bfrac = X
            Case Is < 240
                rfrac = 0
                gfrac = X
                bfrac = c
            Case Is < 300
                rfrac = X
                gfrac = 0
                bfrac = c
            Case Else
                rfrac = c
                gfrac = 0
                bfrac = X
        End Select
        r = Round((rfrac + m) * 255)
        g = Round((gfrac + m) * 255)
        b = Round((bfrac + m) * 255)
    End If
    
    Dim rgbArr(2) As Integer
    rgbArr(0) = r
    rgbArr(1) = g
    rgbArr(2) = b
    
    HSLtoRGB = rgbArr
End Function
