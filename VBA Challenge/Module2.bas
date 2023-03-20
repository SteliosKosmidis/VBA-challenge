Attribute VB_Name = "Module2"
Sub Greatest()

'go through all worksheets
For Each ws In Worksheets

'Determine last row
LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  'set headers
  ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
            ' Start Loop For Final Results
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
       
    Next ws

End Sub

