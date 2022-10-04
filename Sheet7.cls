VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Runs when clicking a new cell - The chosen cell should greater than B2 and the starting point 's'
'This will run the Sidewinder algorithm, and run DFS on the generated graph
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False
    Call algorithms.generate_SW
    Call algorithms.Driver
    Application.EnableEvents = True
End Sub
