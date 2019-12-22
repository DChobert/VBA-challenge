# VBA-challenge GW Boot Camp VBA Homework

Sub VBAChallenge():

    For Each ws In Worksheets
    
        Dim LastRow As Long
        Dim StockVolume As Double
            StockVolume = 0
        Dim SummaryTableRow As Long
            SummaryTableRow = 2
        Dim StockOpen As Double
        Dim StockClose As Double
        Dim StockDelta As Double
        Dim PreviousAmount As Long
            PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
            GreatestIncrease = 0
        Dim GreatestDecrease As Double
            GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
            GreatestTotalVolume = 0
        Dim StockSymbol As String

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                StockSymbol = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = StockSymbol
                ws.Range("L" & SummaryTableRow).Value = StockVolume
                
                StockVolume = 0

                StockOpen = ws.Range("C" & PreviousAmount)
                StockClose = ws.Range("F" & i)
                StockDelta = StockClose - StockOpen
                ws.Range("J" & SummaryTableRow).Value = StockDelta
                
                    If StockOpen = 0 Then
                        PercentChange = 0
                        Else
                        StockOpen = ws.Range("C" & PreviousAmount)
                        PercentChange = StockDelta / StockOpen
                    End If
                    
                    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                    ws.Range("K" & SummaryTableRow).Value = PercentChange

                    If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                        Else
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    End If
            
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                
                End If
                
            Next i

            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
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
     
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Columns("I:Q").AutoFit

    Next ws

End Sub
