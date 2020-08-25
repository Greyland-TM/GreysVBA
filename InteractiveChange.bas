Attribute VB_Name = "InteractiveChange"


'*****************************************************************************
' ThisCode containes all the functions and modules to updae the sheet.
' Runs when any cell value is changed.
' It is responsible for updating the Boeing list interactively with user input
'*****************************************************************************

Option Explicit

'Displays how many are in the BHI warehouse at the moment

Sub InteractiveQuantities(CheckRow As Integer)
    
    Cells(CheckRow, 9).Value = ""
    Cells(CheckRow, 7).Value = ""
            
    'Debug.Print "CheckRow is: " & CheckRow
    Cells(CheckRow, 9).Value = 0
    Dim j As Integer
    For j = 1 To 1800
        If Worksheets("BHI Stock").Cells(j, 3).Value = Cells(CheckRow, 6).Value Then
            Cells(CheckRow, 9).Value = Cells(CheckRow, 9).Value + Worksheets("BHI Stock").Cells(j, 7).Value
            Cells(CheckRow, 7).Value = Worksheets("BHI Stock").Cells(j, 5).Value
        End If
    Next j
        
    'Checks the CSP warehouse if not found in BHI
    If Not Cells(CheckRow, 9).Value = "" Then
        If Cells(CheckRow, 9).Value = 0 Then
            CheckCSP (CheckRow)
        End If
    End If
    
    If Cells(CheckRow, 9).Value = "" Then
        Cells(CheckRow, 9).Value = 0
        Cells(CheckRow, 7).Value = "Not In Stock"
    End If
End Sub
' Used to calculate how many are needed, hidden information on column K.

Sub InteraactiveTotals(CheckRow As Integer)
    Dim j As Integer
    Cells(CheckRow, 11).Value = Cells(CheckRow, 5).Value
    For j = 5 To CheckRow - 1
        If Not IsEmpty(Cells(CheckRow, 7).Value) Then
            If Cells(CheckRow, 7).Value = Cells(j, 7).Value Then
                Cells(CheckRow, 11).Value = Cells(CheckRow, 11).Value + Cells(j, 5).Value
            End If
        End If
    Next j
End Sub

' Used to calculate how many are needed, hidden information on column L.

Sub InteractiveUpdateStock(CheckRow As Integer)
    'Debug.Print "InteractiveUpdateStock CheckRow: " & CheckRow
    Dim j As Integer
    Cells(CheckRow, 12).Value = Cells(CheckRow, 9).Value
    For j = 5 To CheckRow - 1
        If Not IsEmpty(Cells(CheckRow, 9).Value) And Cells(j, 6).Value = Cells(CheckRow, 6).Value Then
            Cells(CheckRow, 12).Value = Cells(CheckRow, 12).Value - Cells(j, 5).Value
            If Cells(CheckRow, 12).Value < 0 Then
                Cells(CheckRow, 12).Value = 0
            End If
        End If
    Next j
End Sub

'Displays quantity to be ordered from Boeing.
'decides weather a part is added to the list or not.

Sub InteractiveGetNeeded(CheckRow As Integer, OldPN As Variant, OffsetValue As Integer)
    If Not IsEmpty(Cells(CheckRow, 9).Value) Then
        Dim EditRow As Integer
        'Debug.Print "OldPN is: " & OldPN
        EditRow = SeachForPN(CheckRow, OldPN, OffsetValue)
        'Debug.Print "CheckRow is: " & CheckRow
        
    End If
End Sub


' Sets the number code for the lights in column I.

Sub InteractiveSetDots(CheckRow As Integer)
        If Not IsEmpty(Cells(CheckRow, 9).Value) Then
            If Cells(CheckRow, 12).Value < Cells(CheckRow, 5).Value Then
                Cells(CheckRow, 10).Value = 0
            ElseIf IsEmpty(Cells(CheckRow, 9).Value) Then
                Cells(CheckRow, 10).Value = ""
            ElseIf Cells(CheckRow, 12).Value = Cells(CheckRow, 5).Value Then
                Cells(CheckRow, 10).Value = 1
            Else
                Cells(CheckRow, 10).Value = 2
            End If
        End If
End Sub

' Used to display color coded lights.

Sub InteractiveIconSets()
Dim iset As IconSetCondition
Dim rg As Range
Set rg = Range("J5", Range("J200").End(xlDown))
rg.FormatConditions.Delete
Set iset = rg.FormatConditions.AddIconSetCondition
'select the traffic lights iconset
With iset
    .IconSet = ActiveWorkbook.iconsets(xl3TrafficLights1)
    .ReverseOrder = False
    .ShowIconOnly = True
End With
'specify amber traffic light for values >= 80% of target(2500)
With iset.IconCriteria(2)
    .Type = xlConditionValueNumber
    .Value = 1
End With
'specify green traffic light for values >= the target(2500)
With iset.IconCriteria(3)
    .Type = xlConditionValueNumber
    .Value = 2
End With
End Sub

' Displays SWO on parts list
' Only runs if GetNeeded() adds quantitys to parts list.

Sub InteractiveTransferSWO(CheckRow As Integer)
    'Dim CheckRow As Integer
    For CheckRow = 5 To 200
        If Cells(CheckRow, 6).Value = "" Then
            Cells(CheckRow, 14).Value = ""
        Else
            If InStr(Cells(CheckRow, 2).Value, "SWO") > 0 Then
                Cells(CheckRow, 14).Value = Cells(CheckRow, 2).Value
            Else
                Cells(CheckRow, 14).Value = ""
            End If
        End If
    Next CheckRow
End Sub

' Will check if part is already on the list.


Function SeachForPN(initalRow As Integer, OldPN As Variant, OffsetValue As Integer) As Integer
    'Debug.Print "OldPN is: " & OldPN
    
    Dim Found As Boolean
    Found = False

    Dim SWONum As Integer
    SWONum = Cells(initalRow, 13).Value
    
    Dim PNCheck As String
    PNCheck = Cells(initalRow, 6).Value
    
    'Debug.Print "PNCheck Check is: " & PNCheck

    Dim CheckRow As Integer
    CheckRow = 6
    
    Do While Cells(CheckRow, 13).Value <= SWONum + 1
        'Debug.Print "Checking: " & Cells(CheckRow, 15).Value
        'Debug.Print "OldPN is: " & OldPN
        If Cells(initalRow, 6).Value = "" Then
            Exit Do
        End If
    
        If Cells(CheckRow, 13).Value = SWONum Then
        
            If Cells(CheckRow, 15).Value = PNCheck Then
                Cells(CheckRow, 17).Value = Cells(initalRow, 5).Value - Cells(initalRow, 12).Value
                Found = True
                
                If Cells(CheckRow, 17).Value <= 0 Then
                    Cells(CheckRow, 15).Value = ""
                    Cells(CheckRow, 16).Value = ""
                    Cells(CheckRow, 17).Value = ""
                    Call InteractiveShiftList(SWONum, CheckRow, OffsetValue)
                    CheckRow = CheckRow + 1

                Else
                    CheckRow = CheckRow + 1
                End If
            
            ElseIf Cells(CheckRow, 15).Value = OldPN Then
                    Cells(CheckRow, 15).Value = ""
                    Cells(CheckRow, 16).Value = ""
                    Cells(CheckRow, 17).Value = ""
                    Call InteractiveShiftList(SWONum, CheckRow, OffsetValue)
                    CheckRow = CheckRow + 1
            End If
        End If

        If SWONum < Cells(CheckRow, 13).Value And Found = False Then
            If Cells(initalRow, 5).Value - Cells(initalRow, 12).Value > 0 Then
                CheckRow = CheckRow - 1
                'Debug.Print "InitialRow is: " & initalRow & " & " & "CheckRow is: " & CheckRow & " & " & "SWONum is: " & SWONum
                Call AppendNewRow(initalRow, CheckRow, SWONum)
            End If
            Exit Do
        End If

        If Cells(CheckRow, 13).Value = "" Then
            Exit Do
        End If
        
        CheckRow = CheckRow + 1
    Loop
End Function
'Finds the bottom of the list and appends the new item

Sub AppendNewRow(AppendingRow As Integer, InsertionRow As Integer, SWONum As Integer)
    
    InsertionRow = InsertionRow - 1
    
    Dim OrigSWONum As Integer
    OrigSWONum = SWONum
    
    Do While OrigSWONum = SWONum
    
        SWONum = Cells(AppendingRow, 13).Value
        
        If Cells(InsertionRow, 15).Value = "" Then
            InsertionRow = InsertionRow - 1
        ElseIf Cells(AppendingRow, 5).Value - Cells(AppendingRow, 12).Value <= 0 Then
            Cells(InsertionRow, 15).Value = ""
            Cells(InsertionRow, 16).Value = ""
            Cells(InsertionRow, 17).Value = ""
            Exit Do
        Else
            InsertionRow = InsertionRow + 1
            Cells(InsertionRow, 15).Value = Cells(AppendingRow, 6).Value
            Cells(InsertionRow, 16).Value = Cells(AppendingRow, 7).Value
            Cells(InsertionRow, 17).Value = Cells(AppendingRow, 5).Value - Cells(AppendingRow, 12).Value
            Exit Do
        End If
        If Cells(InsertionRow, 15).Value = "All Parts Available." Then
            If Cells(AppendingRow, 5).Value - Cells(AppendingRow, 12).Value > 0 Then
                Cells(InsertionRow, 15).Value = Cells(AppendingRow, 6).Value
                Cells(InsertionRow, 16).Value = Cells(AppendingRow, 7).Value
                Cells(InsertionRow, 17).Value = Cells(AppendingRow, 5).Value - Cells(AppendingRow, 12).Value
                Exit Do
            End If
        End If
        
    Loop
End Sub

' Shifts the Boeing list up to keep organized

Sub InteractiveShiftList(SWONum As Integer, ChangedRow As Integer, OffsetValue As Integer)
    ChangedRow = ChangedRow + 1
    If Not Cells(ChangedRow, 15).Value = "" Then
        Do While Cells(ChangedRow, 13).Value = SWONum
            If OffsetValue = -1 Then
            Cells(ChangedRow, 15).Offset(OffsetValue, 0).Value = Cells(ChangedRow, 15).Value
            Cells(ChangedRow, 16).Offset(OffsetValue, 0).Value = Cells(ChangedRow, 16).Value
            Cells(ChangedRow, 17).Offset(OffsetValue, 0).Value = Cells(ChangedRow, 17).Value
            Cells(ChangedRow, 15).ClearContents
            Cells(ChangedRow, 16).ClearContents
            Cells(ChangedRow, 17).ClearContents
            ChangedRow = ChangedRow + 1
            Else: Exit Do
            End If
        Loop
    End If
End Sub

' Looks through the other short lists for the same PN

Sub SearchForOthers(CheckRow As Integer, CheckNSN As String, OldPN As Variant)

    Dim EmptyRow As Integer
    EmptyRow = 0
    
    CheckRow = CheckRow + 1

    Do While EmptyRow < 2
    
        If Cells(CheckRow, 6).Value = "" Then
            EmptyRow = EmptyRow + 1
            CheckRow = CheckRow + 1
        ElseIf Cells(CheckRow, 6).Value = CheckNSN Then
            InteractiveQuantities (CheckRow)
            InteraactiveTotals (CheckRow)
            InteractiveUpdateStock (CheckRow)
            Call InteractiveGetNeeded(CheckRow, OldPN, -1)
            InteractiveSetDots (CheckRow)
            InteractiveIconSets
            PrintAllParts
            EmptyRow = 0
            CheckRow = CheckRow + 1
        Else
            EmptyRow = 0
            CheckRow = CheckRow + 1
        End If
        
    Loop
End Sub

