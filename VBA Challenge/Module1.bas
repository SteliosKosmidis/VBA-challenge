Attribute VB_Name = "Module1"
Sub Stock():


For Each ws In Worksheets

        ' Set headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Set Variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
       

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            


                ' Set Ticker Name
                TickerName = ws.Cells(i, 1).Value
                ' Print The Ticker Name In The Summary Table
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ' Print The Ticker Total Amount To The Summary Table
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                ' Reset Ticker Total
                TotalTickerVolume = 0

' Add To Ticker Total Volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            ' Check If We Are Still Within The Same Ticker Name If It Is Not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                ' Determine Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' Format Double To Include % Symbol And Two Decimal Places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add One To The Summary Table Row
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i



            '

    Next ws

End Sub
