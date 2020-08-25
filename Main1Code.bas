Attribute VB_Name = "Main1Code"


'*********************************************************
' This code runs once Button 1 is pressed for Boeing list.
' It is responsible for updating the whole sheet at once.
'*********************************************************

Option Explicit

'emptys columns to create a new list

Sub ClearColumns()
    Range("G6:Q199").ClearContents
End Sub

'New Sub for checking GVT-01 item under Boeing list
'Displays how many are in the GVT-01 warehouse instead of BHI

Sub GetQuantitysGVT(ByVal CheckRow As Integer, ByVal NSN As String)
    Dim j As Integer
    Cells(CheckRow, 9).Value = 0
    
    For j = 1 To 9000
    Dim ActualNSN As String
        ActualNSN = Replace(NSN, "(GVT-01)", "")
        If Worksheets("GVT-01 Stock").Cells(j, 3).Value = ActualNSN Then
            Cells(CheckRow, 9).Value = Cells(CheckRow, 9).Value + Worksheets("GVT-01 Stock").Cells(j, 7).Value
            If Cells(CheckRow, 7).Value = "" Then
                Cells(CheckRow, 7).Value = Worksheets("GVT-01 Stock").Cells(j, 5).Value
            End If
        End If
    Next j
End Sub

'Displays how many are in the BHI warehouse at the moment

Sub GetQuantitys(CheckRow As Integer)
    If Not Cells(CheckRow, 6).Value = "" Then
        Cells(CheckRow, 9).Value = 0
        Dim j As Integer
        For j = 1 To 1800
            If Worksheets("BHI Stock").Cells(j, 3).Value = Cells(CheckRow, 6).Value Then
                Cells(CheckRow, 9).Value = Cells(CheckRow, 9).Value + Worksheets("BHI Stock").Cells(j, 7).Value
                    
                If Cells(CheckRow, 7).Value = "" Then
                    Cells(CheckRow, 7).Value = Worksheets("BHI Stock").Cells(j, 5).Value
                End If
            End If
        Next j
    End If
    'Checks the CSP warehouse if not found in BHI
    If Not Cells(CheckRow, 9).Value = "" Then
        If Cells(CheckRow, 9).Value = 0 Then
            CheckCSP (CheckRow)
        End If
    End If
End Sub

' Used to calculate how many are needed, hidden information on column K.

Sub Totals(ByVal CheckRow As Integer)
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

Sub UpdateInStock(ByVal CheckRow As Integer)
    Dim j As Integer
    Cells(CheckRow, 12).Value = Cells(CheckRow, 9).Value
    For j = 5 To CheckRow - 1
        If Not IsEmpty(Cells(CheckRow, 9).Value) And Cells(j, 6).Value = Cells(CheckRow, 6).Value Then
            Cells(CheckRow, 12).Value = Cells(CheckRow, 12).Value - Cells(j, 5).Value
            If Cells(CheckRow, 12).Value < 0 Then
                Cells(CheckRow, 12).Value = 0
            End If
        Else: Cells(CheckRow, 12).Value = Cells(CheckRow, 9).Value
        End If
    Next j
End Sub

'Displays quantity to be ordered from Boeing.
'decides weather a part is added to the list or not.

Sub GetNeeded(ByVal CheckRow As Integer)
    If Not IsEmpty(Cells(CheckRow, 9).Value) Then
        If Cells(CheckRow, 12).Value - Cells(CheckRow, 5).Value > 0 Then
            Cells(CheckRow, 17).Value = ""
        ElseIf Cells(CheckRow, 12).Value = Cells(CheckRow, 5).Value Then
            Cells(CheckRow, 17).Value = ""
        Else
            Cells(CheckRow, 17).Value = Abs(Cells(CheckRow, 12).Value - Cells(CheckRow, 5).Value)
        End If
    End If
End Sub

' Sets the number code for the lights in column I.

Sub SetDots()
    Dim CheckRow As Integer
    For CheckRow = 5 To 200
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
    Next CheckRow
End Sub

' Used to display color coded lights.

Sub iconsets()
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

Sub TransferSWO()
    Dim CheckRow As Integer
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

' Displays Nomenclature on parts list
' Only runs if GetNeeded() adds quantitys to parts list.

Sub TransferNomenclature()
    Dim CheckRow As Integer
        For CheckRow = 5 To 200
            If Cells(CheckRow, 17).Value = "" Then
                Cells(CheckRow, 15).Value = ""
            Else
                Cells(CheckRow, 15).Value = Cells(CheckRow, 6).Value
            End If
        Next CheckRow
End Sub

' Displays PN on parts list
' Only runs if GetNeeded() adds quantitys to parts list.

Sub TransferPN()
    Dim CheckRow As Integer
        For CheckRow = 5 To 200
            If Cells(CheckRow, 17).Value = "" Then
                Cells(CheckRow, 16).Value = ""
            Else
                Cells(CheckRow, 16).Value = Cells(CheckRow, 7).Value
            End If
        Next CheckRow
End Sub

' Gives each new swo a number that displays on every row.
' Used to colapse the list

Sub NumberSWOs()
    
    Dim BlankRows As Integer
    BlankRows = 0
    
    Dim CheckRow As Integer
    Dim count As Integer
    
    For CheckRow = 5 To 200
        If Not BlankRows > 10 Then
            If InStr(Cells(CheckRow, 2).Value, "SWO") > 0 Then
                count = count + 1
                Cells(CheckRow, 13).Value = count
            Else
                If Cells(CheckRow, 6).Value = "" Then
                    BlankRows = BlankRows + 1
                    Cells(CheckRow, 13).Value = count
                Else
                    BlankRows = 0
                    Cells(CheckRow, 13).Value = count
                End If
            End If
        End If
    Next CheckRow
End Sub

' colapses the list so it it organized

Sub ShiftList()
    Dim CheckRow As Integer
    Dim MoveUP As Integer
    Dim off As Integer
    Dim check As Boolean
    
    For CheckRow = 5 To 200
    
    check = True
    off = 0
        If Not IsEmpty(Cells(CheckRow, 15).Value) Then
            MoveUP = CheckRow
            While check
                If MoveUP = CheckRow Then
                    MoveUP = MoveUP - 1
                ElseIf Cells(MoveUP, 15).Value = "" And Cells(CheckRow, 13).Value = Cells(MoveUP, 13).Value Then
                    MoveUP = MoveUP - 1
                    off = off - 1
                Else
                    check = False
                End If
            Wend
            If off < 0 Then
                Cells(CheckRow, 15).Offset(off, 0).Value = Cells(CheckRow, 15).Value
                Cells(CheckRow, 16).Offset(off, 0).Value = Cells(CheckRow, 16).Value
                Cells(CheckRow, 17).Offset(off, 0).Value = Cells(CheckRow, 17).Value
                Cells(CheckRow, 15).ClearContents
                Cells(CheckRow, 16).ClearContents
                Cells(CheckRow, 17).ClearContents
            End If
            off = 0
        End If
    Next CheckRow
End Sub

' Displays "All Parts Availabe" if no list was created.

Sub PrintAllParts()
    Dim CheckRow As Integer
    For CheckRow = 5 To 200
        If Not Cells(CheckRow, 14).Value = "" Then
            If Cells(CheckRow, 15).Value = "" Then
                Cells(CheckRow, 15).Value = "All Parts Available."
            End If
        End If
    Next CheckRow
End Sub

' Checks CSP warehouse for part.
' Called in GetQuantitys Sub

Sub CheckCSP(ByVal CellToCheck As Integer)
    Dim CSPStock As Integer
    CSPStock = 0
    Dim j As Integer
    For j = 1 To 900
        If Worksheets("CSP").Cells(j, 3).Value = Cells(CellToCheck, 6).Value Then
            CSPStock = Worksheets("CSP").Cells(j, 7).Value
        End If
    Next j
    If CSPStock > 0 Then
        Dim MSG As String
        MSG = "(" & CSPStock & " CSP)"
        Cells(CellToCheck, 7).Value = Cells(CellToCheck, 7).Value & MSG & " "
    End If
End Sub

