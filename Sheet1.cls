VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'***********************************************************************************
'This is the Main Code to be run to update the spreadsheet as required.
'It calls all of the other necessary Main Functions.
'It will run only when a certian cells are activated by selecting or chaning values.
'***********************************************************************************

Option Explicit

Dim OldValue As Variant
Dim CurrentLocation As Integer
Dim CurrentColumn As Integer


' Button Avtivation

Private Sub CommandButton1_Click()
    Main1
End Sub

' Button Avtivation

Private Sub CommandButton2_Click()
    Main2
End Sub

Sub Main1()
    Call RunMain1
End Sub

Sub Main2()
    Call RunMain2
End Sub

' Saves the original value of selected cell

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim intersection As Range
    Set intersection = Intersect(Target, Range("E6:F200"))

    If Not intersection Is Nothing Then
        OldValue = Target.Value
        CurrentLocation = Target.Row()
        CurrentColumn = Target.Column()
        Debug.Print "CELL SELECTION: " & "OldValue is: " & OldValue & " & " & "CurrentLocation is: " & CurrentLocation & " & " & " CurrentColumn is: " & CurrentColumn
    End If
End Sub

'Updates the sheet when a cell is changed

Public Sub Worksheet_Change(ByVal Target As Range)
    Dim intersection As Range
    Set intersection = Intersect(Target, Range("E6:F200"))

    If Not intersection Is Nothing Then
    
        Debug.Print "CELL CHANGE: " & "OldValue is: " & OldValue & " & " & "CurrentLocation is: " & CurrentLocation & " & " & " CurrentColumn is: " & CurrentColumn
    
        Dim CheckRow As Integer
        CheckRow = Target.Row()
        
        Dim SWONum As Integer
        SWONum = Cells(CheckRow, 13).Value
        
        ActiveSheet.Unprotect "greycy9391"
        Application.EnableEvents = False
        
        Debug.Print "OldValue is now: " & OldValue
        OldValue = CellsChanged(CheckRow, SWONum, CurrentLocation, OldValue, CurrentColumn)
        CurrentLocation = CheckRow
        Debug.Print "OldValue is now: " & OldValue
        Debug.Print "Final Check: " & "OldValue is: " & OldValue & " & " & "CurrentLocation is: " & CurrentLocation & " & " & " CurrentColumn is: " & CurrentColumn
        
        ActiveSheet.Protect "greycy9391", True, True
        Application.EnableEvents = True

    End If
    
End Sub
    
