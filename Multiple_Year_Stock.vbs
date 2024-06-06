Sub Multiple_Year_Stock()

'Skips displaying the entire process and only displays final results
Application.ScreenUpdating = False

'For Loop for each worksheet
For Each ws In Worksheets

    Dim row_count As Integer
    Dim Ticker_symbol As String
    Dim Open_price As Double
    Dim Close_price As Double
    Dim Open_price_Tracker As Boolean
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Stock_volume As LongLong
    
    Dim Max As Double
    Dim Min As Double
    Dim Max_Vol As LongLong
   
    

    'Creating Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("I1:Q1").Font.Bold = True
        
    'Set values for LastRow, Booleen to find the open price, and row_count to display results
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Open_price_Tracker = True
    row_count = 2
    
        'Start For Loop to find Tickers, Yearly Change, Percentage Change, and Total Stock Volume
         For i = 2 To LastRow
    
            'Compares the Tickers in the current row to the next.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Sets Ticker Symbol and equations
                Ticker_symbol = ws.Cells(i, 1).Value
                Yearly_Change = Close_price - Open_price
                Percent_change = Yearly_Change / Open_price
                Stock_volume = Stock_volume + ws.Cells(i, 7).Value
            
                'Input results from above in appropriate cells
                ws.Range("I" & row_count).Value = Ticker_symbol
                ws.Range("J" & row_count).Value = Yearly_Change
                ws.Range("J" & row_count).NumberFormat = "$0.00"
                ws.Range("K" & row_count).Value = Percent_change
                ws.Range("K" & row_count).NumberFormat = "0.00%"
                ws.Range("L" & row_count).Value = Stock_volume
           
            
                'add to row count
                row_count = row_count + 1
                
                'rests stock volume total
                Stock_volume = 0
                
                'rest Open_price_tracker to True
                Open_price_Tracker = True
            
            
             Else
                ' Conditional Statement for Open Price. Resets boolean to false after it gets the open price
                If Open_price_Tracker = True Then
                    Open_price = ws.Cells(i, 3)
                    Open_price_Tracker = False
                    Stock_volume = Stock_volume + ws.Cells(i, 7).Value
                ' Sets closing price
                Else
                    Close_price = ws.Cells(i + 1, 6).Value
                    Stock_volume = Stock_volume + ws.Cells(i, 7).Value
             End If
            End If
        Next i
        
        
        'Find the last row for Column K
        LastRow_K = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Conditional formatting for column K (Percentage Change)
        For i = 2 To LastRow_K
            If ws.Cells(i, 10).Value <= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
        Next i
        
        
        
        'Creating fields for next section
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O2:O4").Font.Bold = True
        
        
        'Ensure to find the Max/Min for all 3 using Max/Min function from worksheet functions
        Max = WorksheetFunction.Max(ws.Range("K:K"))
        Min = WorksheetFunction.Min(ws.Range("K:K"))
        Max_Vol = WorksheetFunction.Max(ws.Range("L:L"))
        
        
        'Goes through all cells with value in column K and L to return Tickers with Max/Min Value
        For i = 2 To LastRow_K
        
            
            'Finds Max Percentage Change Ticker
            If ws.Cells(i, 11) = Max Then
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            
            'Finds Min Percentage Change Ticker
            If ws.Cells(i, 11).Value = Min Then
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
                
                
            ' Find Max Total Volume Ticker
            If ws.Cells(i, 12) = Max_Vol Then
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
            End If
        Next i
       
    'Formats column to fit all values
    ws.Columns("I:L").AutoFit
    ws.Columns("O:Q").AutoFit
    
Next ws
            
    

End Sub



