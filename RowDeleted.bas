Attribute VB_Name = "RowDeleted"


'****************************************************************************
' This code runs when a row is deleted.
' It will delete the row, the item from the list, and sfit the whol sheet up.
'****************************************************************************

Sub ClearEmptyRow(OrigDeleteRow As Integer, DeletePN As Variant)

    Application.EnableEvents = False
    Dim DeleteRow As Integer
    DeleteRow = OrigDeleteRow
    Dim SWONum As Integer
    SWONum = Cells(DeleteRow, 13).Value
    
    Dim FindStart As Integer
    FindStart = Cells(DeleteRow, 13).Value
    Cells(DeleteRow, 5).ClearContents
    Cells(DeleteRow, 7).ClearContents
    Cells(DeleteRow, 8).ClearContents
    Cells(DeleteRow, 9).ClearContents
    Cells(DeleteRow, 10).ClearContents
 
    Do While FindStart >= SWONum
        DeleteRow = DeleteRow - 1
        FindStart = Cells(DeleteRow, 13).Value
    Loop
    
    DeleteRow = DeleteRow + 1
    Debug.Print "SWONum: " & SWONum
    Debug.Print "FindStart: " & FindStart
    Debug.Print "Cells(DeleteRow, 13).Value: " & Cells(DeleteRow, 13).Value
    Debug.Print "Cells(FindStart, 13).Value: " & Cells(FindStart, 13).Value
    
    ' **** #6 Problem is here with Cells(DeleteRow, 13).Value****
    
    Do While Cells(FindStart, 13).Value <= SWONum + 1
        If Cells(DeleteRow, 13).Value = SWONum Then
            If Cells(DeleteRow, 15).Value = DeletePN Then
                Cells(DeleteRow, 11).ClearContents
                Cells(DeleteRow, 12).ClearContents
                Cells(DeleteRow, 15).ClearContents
                Cells(DeleteRow, 16).ClearContents
                Cells(DeleteRow, 17).ClearContents
            Exit Do
            Else
                DeleteRow = DeleteRow + 1
            End If
        Else
            DeleteRow = DeleteRow + 1
        End If
        
        If SWONum < Cells(DeleteRow, 13).Value Then
            Exit Do
        End If
    Loop
    
    Application.EnableEvents = True
    Call InteractiveShiftList(SWONum, DeleteRow, 1)
    
End Sub

Sub UpdateOthers(StartingRow As Integer, SWONum As Integer)
'Sub DeleteAndUpdateRow(StartingRow As Integer, OffsetValue As Integer, SWONum As Integer)
    Application.EnableEvents = False
    
    Dim BlankRow As Integer
    BlankRow = 0
    StartingRow = StartingRow + 1
    
    Do While BlankRow < 2
        If Cells(StartingRow, 6).Value = "" Then
            Call DeleteAndUpdateRow(StartingRow, -1, SWONum)
            StartingRow = StartingRow + 1
            BlankRow = BlankRow + 1
        Else
            Call DeleteAndUpdateRow(StartingRow, -1, SWONum)
            StartingRow = StartingRow + 1
            BlankRow = 0
        End If
    Loop

    Application.EnableEvents = True
    
End Sub
