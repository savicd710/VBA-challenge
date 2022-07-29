Sub stock_tracker()

    'To loop the macro through all worksheets
    For each ws In Worksheets

        'Create the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"                                               
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Create and set the variables
        Dim Ticker as String
        Dim TotalVolume as Double
        TotalVolume = 0

        Dim YearOpen as Double
        Dim YearClose as Double
        Dim YearChange as Double
        Dim PercentChange as Double

        Dim PreviousAmt as Long
        PreviousAmt = 2
        Dim SummaryRow as Long
        SummaryRow = 2

        Dim GreatestInc as Double
        Dim GreatestDec as Double
        Dim GreatestTotVol as Double
        GreatestInc = 0
        GreatestDec = 0
        GreatestTotVol = 0

        'Last row calculation
        Dim LastRow as Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'creating the summary table
        For i = 2 To LastRow

            'adding ticker volumes together
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            'have to make sure the ticker is still the same using if statement
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'set ticker and print ticker/volume in the summary table
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryRow).Value = Ticker
                ws.Range("L" & SummaryRow).Value = TotalVolume

                'make sure to reset volume for next stock
                TotalVolume = 0

                'calculate yearly open, close, and change
                YearOpen = ws.Range("C" & PreviousAmt)
                YearClose = ws.Range("F" & i)
                YearChange = YearClose - YearOpen
                ws.Range("J" & SummaryRow).Value = YearChange

                'use the (new-old)/old formula to calculate percent change
                If YearOpen = 0 Then
                    PercentChange = 0
                Else
                    YearOpen = ws.Range("C" & PreviousAmt)
                    PercentChange = YearChange / YearOpen
                End If

                'format the percent change column to have % and 2 decimals
                ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryRow).Value = PercentChange

                'conditional formatting for positive (green) and negative (red) change
                If ws.Range("J" & SummaryRow).Value >= 0 Then
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                End If

                '+1 to summary row
                SummaryRow = SummaryRow + 1
                PreviousAmt = i + 1
                End If
            Next i

            ' greatest % Inc/Dec and Vol
            LastRow = ws.Cells(Rows.Count, 11)End(xlUp).Row

            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value > ws.Range("K" & i).Value
                    ws.Range("P2").Value > ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value > ws.Range("Q3").Value Then
                    ws.Range("Q3").Value > ws.Range("K" & i).Value
                    ws.Range("P3").Value > ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value > ws.Range("L" & i).Value
                    ws.Range("P4").Value > ws.Range("I" & i).Value
                End If

            Next i 

            'format the percent change column to have % and 2 decimals
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
        
        'format table columns
        ws.Columns("I:Q").AutoFit

    'finally close the for loop at the beginning of script to go to next ws
    Next ws

End Sub


