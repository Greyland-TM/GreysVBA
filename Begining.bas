Attribute VB_Name = "Begining"




'*******************************************************************
' This is the main code blocts that run once a cell has been changed
' Run after Boing update button is pressed or a cell is changed
' It calls all of the other necessary functions and modules.
'*******************************************************************

Sub RunMain1()
    ActiveSheet.Unprotect "greycy9391"
    ClearColumns
    Dim CheckRow As Integer
    For CheckRow = 5 To 200
        If Not Cells(CheckRow, 6).Value = "" Then
            If InStr(Cells(CheckRow, 6).Value, "GVT-01") > 0 Then
                Dim NSN As String
                NSN = Cells(CheckRow, 6).Value
                Call GetQuantitysGVT(CheckRow, NSN)
                Totals (CheckRow)
                UpdateInStock (CheckRow)
                GetNeeded (CheckRow)
            Else
                GetQuantitys (CheckRow)
                Totals (CheckRow)
                UpdateInStock (CheckRow)
                GetNeeded (CheckRow)
            End If
        End If
    Next CheckRow
    TransferSWO
    SetDots
    iconsets
    TransferNomenclature
    TransferPN
    NumberSWOs
    ShiftList
    PrintAllParts
    ActiveSheet.Protect "greycy9391", True, True
End Sub

' Run after GVT-01 update button is pressed
' Code in Module2

Sub RunMain2()
    ActiveSheet.Unprotect "greycy9391"
    ClearColumns2
    GetQuantitys2
    Totals2
    UpdateInStock2
    GetNeeded2
    SetDots2
    iconsets2
    TransferSWO2
    TransferNomenclature2
    TransferPN2
    NumberSWOs2
    ShiftList2
    PrintAllParts2
    ActiveSheet.Protect "greycy9391", True, True
End Sub


Function CellsChanged(CheckRow As Integer, SWONum As Integer, CurrentLocation As Integer, OldValue As Variant, CurrentColumn As Integer) As Variant

    If Not Cells(CheckRow, 5).Value = "" And Not Cells(CheckRow, 6).Value = "" Then

        Dim NSN As String
        NSN = Cells(CheckRow, 6).Value

        If InStr(Cells(CheckRow, 6).Value, "GVT-01") > 0 Then
            Call GetQuantitysGVT(CheckRow, NSN)
            Totals (CheckRow)
            UpdateInStock (CheckRow)
            GetNeeded (CheckRow)
            InteractiveSetDots (CheckRow)
            InteractiveIconSets
            PrintAllParts
            'Call SearchForOthers(CheckRow, NSN)
            'Call UpdateGVT
        Else
            InteractiveQuantities (CheckRow)
            InteraactiveTotals (CheckRow)
            InteractiveUpdateStock (CheckRow)
            Call InteractiveGetNeeded(CheckRow, OldValue, -1)
            InteractiveSetDots (CheckRow)
            InteractiveIconSets
            PrintAllParts
            Call SearchForOthers(CheckRow, NSN, OldValue)
            CellsChanged = "Item Changed"
            
            'Call UpdateBHI
        End If
    Else
        If Cells(CheckRow, 6).Value = "" And Not OldValue = "" Then
            Call ClearEmptyRow(CurrentLocation, OldValue)
            Call UpdateOthers(CurrentLocation, SWONum)
            CellsChanged = Cells(CheckRow, 6).Value
        End If
    End If

    If OldValue = "" And Not Cells(CheckRow, 6).Value = "" Then
        If CurrentColumn = 6 Then
            'Debug.Print "CurrentLocation is: " & CurrentLocation
            Call AddedItem(CurrentLocation, SWONum)
            CellsChanged = ""
        End If
    End If
    
    If InStr(Cells(CheckRow, 2).Value, "SWO") > 0 Then
        Cells(CheckRow, 14).Value = Cells(CheckRow, 2).Value
    End If
    
    
    'CellsChanged = "Function Works"
    'Debug.Print "CellsChanged is: " & CellsChanged

End Function
