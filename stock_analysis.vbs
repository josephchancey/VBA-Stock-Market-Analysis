' Stock Analysis VBA Script - Joseph Chancey

Sub Stock_Analysis()

'Dim variables to memory
'Integers
Dim days As Integer
Dim j_col As Integer
'Doubles
Dim total As Double 'May make this LongLong if issues arise, some stocks have very large volumes
Dim change As Double
Dim Percent_Change As Double
Dim averageChange As Double
'Longs (LongLong is used due to length of volume, Long throws errors, LongLong is required)
Dim i As LongLong
Dim start As LongLong
Dim Row_Count As LongLong


'Create Headers
Range("I1").Value = "Ticker Symbl"
Range("J1").Value = "Yrly Change"
Range("K1").Value = "% Change"
Range("L1").Value = "Total Volume"

'Initialize Default Values
j_col = 0
total = 0
change = 0
'Start at "2" to skip header row
start = 2

Row_Count = Cells(Rows.Count, "A").End(xlUp).Row

' Begin loop through dataset
For i = 2 To Row_Count

    'Update total with cell value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        total = total + Cells(i, 7).Value
        'filtering out 0's
        If total = 0 Then
            Range("I" & 2 + j_col).Value = Cells(i, 1).Value
            Range("J" & 2 + j_col).Value = 0
            Range("K" & 2 + j_col).Value = "%" & 0
            Range("L" & 2 + j_col).Value = 0
        Else 'if not 0
            If Cells(start, 3) = 0 Then
                For find_val = start To i
                    If Cells(find_val, 3).Value <> 0 Then
                        start = find_val
                        Exit For
                    End If
                Next find_val
            End If
        
        'Populating Changes
        change = (Cells(i, 6) - Cells(start, 3))
        Percent_Change = Round((change / Cells(start, 3) * 100), 2)
        start = i + 1
        
        'print vals
        Range("I" & 2 + j_col).Value = Cells(i, 1).Value
        Range("J" & 2 + j_col).Value = Round(change, 2)
        Range("K" & 2 + j_col).Value = "%" & Percent_Change
        Range("L" & 2 + j_col).Value = total
        
        'Color Values depending on our previously-grabbed changes
        If change > 0 Then
            Range("J" & 2 + j_col).Interior.ColorIndex = 4
        ElseIf change < 0 Then
            Range("J" & 2 + j_col).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j_col).Interior.ColorIndex = 0
        End If
        
    End If
    
    'Refresh values back to null for next iteration
    j_col = j_col + 1
    total = 0
    change = 0
    days = 0
    
    Else
        total = total + Cells(i, 7).Value
    End If
    
    
Next i

End Sub
