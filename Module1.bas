Attribute VB_Name = "Module1"

'****************************************************************************
' This will actually remove the necessry lines.
' It is looped through in both "RowAdded and "RowDeleted".
'****************************************************************************

Sub DeleteAndUpdateRow(StartingRow As Integer, OffsetValue As Integer, SWONum As Integer)

'Debug.Print StartingRow & " " & OffsetValue & " " & SWONum & " " & Cells(StartingRow, 13).Value
Dim NSN As String
NSN = Cells(StartingRow, 6).Value

If Cells(StartingRow, 13).Value = SWONum And Not Cells(StartingRow, 6).Value = "" Then
    
    Cells(StartingRow, 5).Offset(OffsetValue, 0).Value = Cells(StartingRow, 5).Value
    Cells(StartingRow, 6).Offset(OffsetValue, 0).Value = Cells(StartingRow, 6).Value
    Cells(StartingRow, 7).Offset(OffsetValue, 0).Value = Cells(StartingRow, 7).Value
    Cells(StartingRow, 13).Offset(OffsetValue, 0).Value = Cells(StartingRow, 13).Value
    Cells(StartingRow, 5).ClearContents
    Cells(StartingRow, 6).ClearContents
    Cells(StartingRow, 7).ClearContents
    Cells(StartingRow, 8).ClearContents
    Cells(StartingRow, 9).ClearContents
    Cells(StartingRow, 10).ClearContents
    Cells(StartingRow, 11).ClearContents
    Cells(StartingRow, 12).ClearContents
    
    If InStr(Cells(StartingRow + OffsetValue, 6).Value, "GVT-01") > 0 Then

    
        Call GetQuantitysGVT(StartingRow + OffsetValue, NSN)
        InteraactiveTotals (StartingRow + OffsetValue)
        InteractiveUpdateStock (StartingRow + OffsetValue)
        Call InteractiveGetNeeded(StartingRow + OffsetValue, OldValue, OffsetValue)
        InteractiveSetDots (StartingRow + OffsetValue)
        InteractiveIconSets
        PrintAllParts
    Else
        InteractiveQuantities (StartingRow + OffsetValue)
        InteraactiveTotals (StartingRow + OffsetValue)
        InteractiveUpdateStock (StartingRow + OffsetValue)
        Call InteractiveGetNeeded(StartingRow + OffsetValue, OldValue, OffsetValue)
        InteractiveSetDots (StartingRow + OffsetValue)
        InteractiveIconSets
        'Call InteractiveShiftList(SWONum, ChangedRow)
        PrintAllParts
    End If

ElseIf Not Cells(StartingRow, 6).Value = "" Then
    If StartingRow = 58 Then
        'Debug.Print "Break"
    End If

    'Debug.Print "OffsetValue is: " & OffsetValue
    Cells(StartingRow, 2).Offset(OffsetValue, 0).Value = Cells(StartingRow, 2).Value
    Cells(StartingRow, 3).Offset(OffsetValue, 0).Value = Cells(StartingRow, 3).Value
    Cells(StartingRow, 4).Offset(OffsetValue, 0).Value = Cells(StartingRow, 4).Value
    Cells(StartingRow, 5).Offset(OffsetValue, 0).Value = Cells(StartingRow, 5).Value
    Cells(StartingRow, 6).Offset(OffsetValue, 0).Value = Cells(StartingRow, 6).Value
    Cells(StartingRow, 7).Offset(OffsetValue, 0).Value = Cells(StartingRow, 7).Value
    Cells(StartingRow, 13).Offset(OffsetValue, 0).Value = Cells(StartingRow, 13).Value
    Cells(StartingRow, 14).Offset(OffsetValue, 0).Value = Cells(StartingRow, 14).Value
    Cells(StartingRow, 2).ClearContents
    Cells(StartingRow, 3).ClearContents
    Cells(StartingRow, 4).ClearContents
    Cells(StartingRow, 5).ClearContents
    Cells(StartingRow, 6).ClearContents
    Cells(StartingRow, 7).ClearContents
    Cells(StartingRow, 8).ClearContents
    Cells(StartingRow, 9).ClearContents
    Cells(StartingRow, 10).ClearContents
    Cells(StartingRow, 11).ClearContents
    Cells(StartingRow, 12).ClearContents
    Cells(StartingRow, 14).ClearContents

    If InStr(Cells(StartingRow + OffsetValue, 6).Value, "GVT-01") > 0 Then
    
        Call GetQuantitysGVT(StartingRow + OffsetValue, NSN)
        InteraactiveTotals (StartingRow + OffsetValue)
        InteractiveUpdateStock (StartingRow + OffsetValue)
        Call InteractiveGetNeeded(StartingRow + OffsetValue, OldValue, OffsetValue)
        InteractiveSetDots (StartingRow + OffsetValue)
        InteractiveIconSets
        PrintAllParts
    Else
        InteractiveQuantities (StartingRow + OffsetValue)
        InteraactiveTotals (StartingRow + OffsetValue)
        InteractiveUpdateStock (StartingRow + OffsetValue)
        Call InteractiveGetNeeded(StartingRow + OffsetValue, OldValue, OffsetValue)
        InteractiveSetDots (StartingRow + OffsetValue)
        InteractiveIconSets
        'Call InteractiveShiftList(SWONum, ChangedRow)
        PrintAllParts
    End If

Else
    'Debug.Print "OffsetValue is: " & OffsetValue
    If StartingRow = 58 Then
        'Debug.Print "Break"
    End If
    Cells(StartingRow, 2).Offset(OffsetValue, 0).Value = Cells(StartingRow, 2).Value
    Cells(StartingRow, 3).Offset(OffsetValue, 0).Value = Cells(StartingRow, 3).Value
    Cells(StartingRow, 4).Offset(OffsetValue, 0).Value = Cells(StartingRow, 4).Value
    Cells(StartingRow, 5).Offset(OffsetValue, 0).Value = Cells(StartingRow, 5).Value
    Cells(StartingRow, 6).Offset(OffsetValue, 0).Value = Cells(StartingRow, 6).Value
    Cells(StartingRow, 7).Offset(OffsetValue, 0).Value = Cells(StartingRow, 7).Value
    Cells(StartingRow, 8).Offset(OffsetValue, 0).Value = Cells(StartingRow, 8).Value
    Cells(StartingRow, 9).Offset(OffsetValue, 0).Value = Cells(StartingRow, 9).Value
    Cells(StartingRow, 10).Offset(OffsetValue, 0).Value = Cells(StartingRow, 10).Value
    Cells(StartingRow, 11).Offset(OffsetValue, 0).Value = Cells(StartingRow, 11).Value
    Cells(StartingRow, 12).Offset(OffsetValue, 0).Value = Cells(StartingRow, 12).Value
    Cells(StartingRow, 13).Offset(OffsetValue, 0).Value = Cells(StartingRow, 13).Value
    Cells(StartingRow, 14).Offset(OffsetValue, 0).Value = Cells(StartingRow, 14).Value
    Cells(StartingRow, 15).Offset(OffsetValue, 0).Value = Cells(StartingRow, 15).Value
    Cells(StartingRow, 16).Offset(OffsetValue, 0).Value = Cells(StartingRow, 16).Value
    Cells(StartingRow, 17).Offset(OffsetValue, 0).Value = Cells(StartingRow, 17).Value
    Cells(StartingRow, 2).ClearContents
    Cells(StartingRow, 3).ClearContents
    Cells(StartingRow, 4).ClearContents
    Cells(StartingRow, 5).ClearContents
    Cells(StartingRow, 6).ClearContents
    Cells(StartingRow, 7).ClearContents
    Cells(StartingRow, 8).ClearContents
    Cells(StartingRow, 9).ClearContents
    Cells(StartingRow, 10).ClearContents
    Cells(StartingRow, 11).ClearContents
    Cells(StartingRow, 12).ClearContents
    Cells(StartingRow, 14).ClearContents
    Cells(StartingRow, 15).ClearContents
    Cells(StartingRow, 16).ClearContents
    Cells(StartingRow, 17).ClearContents

End If
End Sub

