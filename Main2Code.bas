Attribute VB_Name = "Main2Code"


'*********************************************************
' This code runs once Button 2 is pressed for GVT-01 list.
' It is responsible for updating the whole sheet at once.
'*********************************************************

Option Explicit

Sub ClearColumns2()
    Range("AA6:AL199").Clear
End Sub

'Displays how many are in the BHI warehouse at the moment

Sub GetQuantitys2()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 200
        If Not Cells(i, 26).Value = "" Then
            Cells(i, 29).Value = 0
            For j = 1 To 9000
                If Worksheets("GVT-01 Stock").Cells(j, 3).Value = Cells(i, 26).Value Then
                    Worksheets("GVT-01 Stock").Cells(j, 38).Value = "Works"
                    Cells(i, 29).Value = Cells(i, 29).Value + Worksheets("GVT-01 Stock").Cells(j, 7).Value
                    If Cells(i, 27).Value = "" Then
                        Cells(i, 27).Value = Worksheets("GVT-01 Stock").Cells(j, 5).Value
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' Used to calculate how many are needed, hidden information on column K.

Sub Totals2()
    Dim i As Integer
    Dim j As Integer
    For i = 5 To 200
    Cells(i, 31).Value = Cells(i, 28).Value
        For j = 5 To i - 1
            If Not IsEmpty(Cells(i, 27).Value) Then
                If Cells(i, 27).Value = Cells(j, 27).Value Then
                    Cells(i, 31).Value = Cells(i, 31).Value + Cells(j, 28).Value
               End If
            End If
         Next j
    Next i
End Sub

' Used to calculate how many are needed, hidden information on column L.

Sub UpdateInStock2()
    Dim i As Integer
    Dim j As Integer
    For i = 5 To 200
        Cells(i, 32).Value = Cells(i, 29).Value
        For j = 5 To i - 1
            If Not IsEmpty(Cells(i, 29).Value) And Cells(j, 26).Value = Cells(i, 26).Value Then
                Cells(i, 32).Value = Cells(i, 32).Value - Cells(j, 28).Value
                If Cells(i, 32).Value < 0 Then
                    Cells(i, 32).Value = 0
                End If
            End If
        Next j
    Next i
End Sub
            
'Displays quantity to be ordered from Boeing.
'decides weather a part is added to the list or not.
            
Sub GetNeeded2()
    Dim i As Integer
    For i = 5 To 200
        If Not IsEmpty(Cells(i, 29).Value) Then
            If Cells(i, 32).Value - Cells(i, 28).Value > 0 Then
                Cells(i, 37).Value = ""
            ElseIf Cells(i, 32).Value = Cells(i, 28).Value Then
                Cells(i, 37).Value = ""
            Else
                Cells(i, 37).Value = Abs(Cells(i, 32).Value - Cells(i, 28).Value)
            End If
        End If
    Next i
End Sub

' Sets the number code for the lights in column I.

Sub SetDots2()
    Dim i As Integer
    For i = 5 To 200
        If Not IsEmpty(Cells(i, 29).Value) Then
            If Cells(i, 32).Value < Cells(i, 28).Value Then
                Cells(i, 30).Value = 0
            ElseIf IsEmpty(Cells(i, 29).Value) Then
                Cells(i, 30).Value = ""
            ElseIf Cells(i, 32).Value = Cells(i, 28).Value Then
                Cells(i, 30).Value = 1
            Else
                Cells(i, 30).Value = 2
            End If
        End If
    Next i
End Sub

' Used to display color coded lights.

Sub iconsets2()
Dim rg As Range
Dim iset As IconSetCondition
Set rg = Range("AD5", Range("AD200").End(xlDown))
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
' Only runs if GetNeeded2() adds quantitys to parts list.

Sub TransferSWO2()
   Dim i As Integer
    For i = 5 To 200
        If Cells(i, 26).Value = "" Then
            Cells(i, 34).Value = ""
        Else
            If InStr(Cells(i, 22).Value, "SWO") > 0 Then
                Cells(i, 34).Value = Cells(i, 22).Value
            Else
                Cells(i, 34).Value = ""
            End If
        End If
    Next i
End Sub

' Displays Nomenclature on parts list
' Only runs if GetNeeded2() adds quantitys to parts list.

Sub TransferNomenclature2()
    Dim i As Integer
        For i = 5 To 200
            If Cells(i, 37).Value = "" Then
                Cells(i, 35).Value = ""
            Else
                Cells(i, 35).Value = Cells(i, 26).Value
            End If
        Next i
End Sub

' Displays PN on parts list
' Only runs if GetNeeded2() adds quantitys to parts list.

Sub TransferPN2()
    Dim i As Integer
        For i = 5 To 200
            If Cells(i, 37).Value = "" Then
                Cells(i, 36).Value = ""
            Else
                Cells(i, 36).Value = Cells(i, 27).Value
            End If
        Next i
End Sub

' Gives each new swo a number that displays on every row.
' Used to colapse the list

Sub NumberSWOs2()
    Dim i As Integer
    Dim count As Integer
    For i = 5 To 200
        If InStr(Cells(i, 22).Value, "SWO") > 0 Then
            count = count + 1
            Cells(i, 33).Value = count
        Else
            Cells(i, 33).Value = count
        End If
    Next i
End Sub

' colapses the list so it it organized

Sub ShiftList2()
    Dim i As Integer
    Dim MoveUP As Integer
    Dim off As Integer
    Dim check As Boolean
    
    For i = 5 To 200
    
    check = True
    off = 0
        If Not IsEmpty(Cells(i, 35).Value) Then
            MoveUP = i
            While check
                If MoveUP = i Then
                    MoveUP = MoveUP - 1
                ElseIf Cells(MoveUP, 35).Value = "" And Cells(i, 33).Value = Cells(MoveUP, 33).Value Then
                    MoveUP = MoveUP - 1
                    off = off - 1
                Else
                    check = False
                End If
            Wend
            
            If off < 0 Then
                Cells(i, 35).Offset(off, 0).Value = Cells(i, 35).Value
                Cells(i, 36).Offset(off, 0).Value = Cells(i, 36).Value
                Cells(i, 37).Offset(off, 0).Value = Cells(i, 37).Value
                Cells(i, 35).ClearContents
                Cells(i, 36).ClearContents
                Cells(i, 37).ClearContents
            End If
            
            off = 0
        End If
    Next i
End Sub

' Displays "All Parts Availabe" if no list was created.

Sub PrintAllParts2()
    Dim i As Integer
    For i = 5 To 200
        If Not Cells(i, 34).Value = "" Then
            If Cells(i, 35).Value = "" Then
                Cells(i, 35).Value = "All Parts Available."
            End If
        End If
    Next i
End Sub


