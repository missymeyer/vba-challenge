Sub Summarize_Yearly_Stocks():
       
       'Loop through all worksheets
       Dim ws As Worksheet
       For Each ws In Worksheets
    
       'Define column names for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest Percent Increase"
        ws.Range("O3").Value = "Greatest Percent Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        'Declare variables
        Dim TickerName As String
        Dim LastRowA As Long
        Dim LastRowK As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
    
        Dim NewRow As Long
        NewRow = 2
    
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
    
        Dim PrvsAmount As Long
        PrvsAmount = 2
    
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

    'Determine value of the last row by finding the last non-blank cell in column A
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through rows
    For i = 2 To LastRowA
        
        'Add values to total ticker volume
        TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
    
        'Check if the next row has same ticker name as the previous
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set ticker
            TickerName = ws.Cells(i, 1).Value
                
            'Print ticker name
            ws.Range("I" & NewRow).Value = TickerName
                
            'Print ticker volume
            ws.Range("L" & NewRow).Value = TotalTickerVolume
               
            'Reset total ticker volume
            TotalTickerVolume = 0

            'Set open price
            OpenPrice = ws.Range("C" & PrvsAmount)
                
            'and set close price
            ClosePrice = ws.Range("F" & i)
                
            'Cacl change from open to close
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & NewRow).Value = YearlyChange
                
            'Change J to dollars formatting
            ws.Range("J" & NewRow).NumberFormat = "$0.00"

            '% Change with if/then statement
            If OpenPrice = 0 Then
                PercentChange = 0
                    
                'Otherwise, set %Change change/open price
                Else
                YearlyOpen = ws.Range("C" & PrvsAmount)
                PercentChange = YearlyChange / OpenPrice
                        
            End If
                
            'Populate percent change
            ws.Range("K" & NewRow).Value = PercentChange
                
            'Conditional formatting for color
            If ws.Range("J" & NewRow).Value >= 0 Then
            ws.Range("J" & NewRow).Interior.ColorIndex = 4
                    
                Else
                ws.Range("J" & NewRow).Interior.ColorIndex = 3
                
            End If
            
            'Add 1 to rows.
            NewRow = NewRow + 1
              
            PrvsAmount = i + 1
                
        End If
                
        Next i

        'Find last row
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Loop through rows
        For i = 2 To LastRowK
            
            'Greatest percent increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'greatest percent decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Determine volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        'Adjust formatting for percentages
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub