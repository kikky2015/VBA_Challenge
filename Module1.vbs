Sub Yearstock()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        ' Create a Variables
        Dim ticker_name As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim year_delta As Double
        Dim Total_Stock_volume As Double
        Dim Summary_Row As Integer
        Dim percentage_change As Double
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_vol As Double
        Dim max_Inc_Ticker As String
        Dim max_Dec_Ticker As String
        
        ' Set an initial variables for ticker total
        Total_Stock_volume = 0
        max_increase = 0
        max_decrease = 0
        max_vol = 0
        
        ' Keep track of the location for each ticker brand in the summary table
        Summary_Row = 2

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Header rows
        ws.Range("I" & 1).Value = "Ticker"
        ws.Range("J" & 1).Value = "Yearly Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock Volume"
        
        'Initialize opening price for each sheet
            openPrice = ws.Cells(2, 3).Value
            
        ' Loop through all ticker
        For i = 2 To LastRow
        ' Check if we are still within the same ticker name, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Add to the Total Stocks
                Total_Stock_volume = Total_Stock_volume + ws.Cells(i, 7).Value

                ' Print the Ticker name in the Summary Row
                ws.Range("I" & Summary_Row).Value = Cells(i, 1).Value
                
                'Print the Yearly Change in the Summary Row,
                'Closing Price = ws.Cells(i, 6).Value
                year_delta = ws.Cells(i, 6).Value - openPrice
                ws.Range("J" & Summary_Row).Value = year_delta
                
                ' Print the Percentage change in the Summary Row
                If openPrice = 0 Then
                    percentage_change = 0
                Else
                    percentage_change = (year_delta / openPrice)
                End If
                
                ws.Range("K" & Summary_Row).Value = percentage_change
                
                'Format percentage row
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                
                ' Print the Total Stock vol to the Summary Row
                ws.Range("L" & Summary_Row).Value = Total_Stock_volume
                '-----------------------------------------------------------------------------------------------
                ' Color the yearly change cell based on positive or negative value
                If year_delta >= 0 Then
                    ws.Range("J" & Summary_Row).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Range("J" & Summary_Row).Interior.Color = RGB(255, 0, 0)
                End If
                '-----------------------------------------------------------------------------
                '
                'return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
                ws.Range("P" & 1).Value = "Ticker"
                ws.Range("Q" & 1).Value = "Value"
                ws.Range("O" & 2).Value = "Greatest % Increase"
                ws.Range("O" & 3).Value = "Greatest % Decrease"
                ws.Range("O" & 4).Value = "Greatest Total Volume"
                '-----------------------------------------------------------------------------------------
                If percentage_change > max_increase Then
                    max_increase = percentage_change
                    max_Inc_Ticker = ws.Cells(i, 1).Value
                End If
                If percentage_change < max_decrease Then
                    max_decrease = percentage_change
                    max_Dec_Ticker = ws.Cells(i, 1).Value
                End If
                If Total_Stock_volume > max_vol Then
                    max_vol = Total_Stock_volume
                    max_Vol_Ticker = ws.Cells(i, 1).Value
                End If
                
                ' Display max values to the summary table
                ws.Range("P" & 2).Value = max_Inc_Ticker
                ws.Range("Q" & 2).Value = (max_increase * 100) & "%"
                ws.Range("Q" & 2).NumberFormat = "0.00%"
                ws.Range("P" & 3).Value = max_Dec_Ticker
                ws.Range("Q" & 3).Value = (max_decrease * 100) & "%"
                ws.Range("Q" & 3).NumberFormat = "0.00%"
                ws.Range("P" & 4).Value = max_Vol_Ticker
                ws.Range("Q" & 4).Value = max_vol
                
                ' Add one to the summary table row
                Summary_Row = Summary_Row + 1
                
        '----------------------------------------------------------------------------------------------------
                'Reset the ticker summary values for next ticker
                Total_Stock_volume = 0
                ticker_name = ""

                'Next open price
                openPrice = ws.Cells(i + 1, 3).Value
                
                closePrice = 0
                year_delta = 0
                percentage_change = 0
                ' If the cell immediately following a row is the same ticker...
            Else
                ' Add to the  Total Stocks
            
                Total_Stock_volume = Total_Stock_volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Auto-fit columns in the summary table
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
Next ws
End Sub
