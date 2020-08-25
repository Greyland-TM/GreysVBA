Attribute VB_Name = "RowAdded"


'*****************************************************************************
' This code runs when a row is added to the bottom of a list
' It will delete the row, the item from the list, and sfit the whole sheet up.
'*****************************************************************************

Sub AddedItem(AddedRow As Integer, SWONum As Integer)

    Application.EnableEvents = False
    
    Dim EmptyRowCount As Integer
    EmptyRowCount = 0
    
    Dim ListBottom As Integer
    ListBottom = AddedRow

    Do While EmptyRowCount <= 20
        If Cells(ListBottom, 6).Value = "" Then
            EmptyRowCount = EmptyRowCount + 1
            ListBottom = ListBottom + 1
        Else
            ListBottom = ListBottom + 1
            EmptyRowCount = 0
        End If
    Loop
    
    Do While ListBottom > AddedRow
        Call DeleteAndUpdateRow(ListBottom, 1, SWONum)
        Cells(ListBottom, 15).Offset(1, 0).Value = Cells(ListBottom, 15).Value
        Cells(ListBottom, 16).Offset(1, 0).Value = Cells(ListBottom, 16).Value
        Cells(ListBottom, 17).Offset(1, 0).Value = Cells(ListBottom, 17).Value
        Cells(ListBottom, 15).ClearContents
        Cells(ListBottom, 16).ClearContents
        Cells(ListBottom, 17).ClearContents
        ListBottom = ListBottom - 1
    Loop
    
    Cells(AddedRow + 1, 13).Value = Cells(AddedRow, 13).Value
    Application.EnableEvents = True
    
End Sub
