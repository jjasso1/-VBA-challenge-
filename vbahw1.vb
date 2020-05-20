Sub vbatesting():
    Dim total As Double
    Dim ws As Worksheet
    Dim j As Integer
    
    For Each ws In Worksheets
    total = 0
    j = 0

    ' Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total = total + ws.Cells(i, 7).Value
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = total
            total = 0
            j = j + 1
        
        Else
            total = total + ws.Cells(i, 7).Value
        End If
    Next i
    total = 0
    j = 0
    Next ws
    
    


End Sub
