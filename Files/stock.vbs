Sub stock()

'Loop through each year worksheets
'Loop through each stock data in each worksheet
'yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.
'The total Volume of the stock
'Ticker symbol
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".



'Declare variables
Dim TotalVolume as Double
Dim StartingPrice as Double
Dim LastRow as Long
Dim TickerCount as Integer
Dim GreatestIncrease as Integer
Dim GreatestDecrease as Integer
Dim GreatestTotal as Integer
For Each ws in Worksheets
    'Set Last Row in Worksheet and Reset Ticker Count and Greatests
    LastRow = ws.Cells(Rows.Count,1).End(xlUp).Row
    TickerCount = 2
    StartingPrice = ws.cells(2,3).Value
    GreatestIncrease = 3
    GreatestDecrease = 3
    GreatestTotal = 3
    'Fill in Headers for tables
    ws.Cells(1,9).Value = "Ticker"
    ws.Cells(1,10).Value = "Yearly Change"
    ws.Cells(1,11).Value = "Percent Change"
    ws.Cells(1,12).Value = "Total Stock Volume"
    ws.Cells(1,16).Value = "Ticker"
    ws.Cells(1,17).Value = "Value"
    ws.Cells(2,15).Value = "Greatest % Increase"
    ws.Cells(3,15).Value = "Greatest % Decrease"
    ws.Cells(4,15).Value = "Greatest Total Volume"

    For i = 2 to LastRow
        'Add current volume to TotalVolume
        TotalVolume = TotalVolume + ws.Cells(i,7).Value
        'Run through tickers until they're different then fill in info and reset counters
        If ws.Cells(i,1).Value <> ws.Cells(i+1,1).Value Then
            'Insert Ticker
            ws.Cells(TickerCount,9).Value = ws.Cells(i,1).Value
            'Insert YearlyChange Formatted
            ws.Cells(TickerCount,10).Value = ws.Cells(i,6).Value - StartingPrice
            ws.Cells(TickerCount,10).NumberFormat = "0.00"
            If ws.Cells(TickerCount,10).Value > 0 Then
                ws.Cells(TickerCount,10).Interior.ColorIndex = 4
            Else
                ws.Cells(TickerCount,10).Interior.ColorIndex = 3
            End If
            'Insert PercentChange with Case for 0 and Format
            If StartingPrice = 0 then
                ws.Cells(TickerCount,11).Value = 1
            Else
                ws.Cells(TickerCount,11).Value = ws.Cells(TickerCount,10).Value/StartingPrice
            End If
                ws.Cells(TickerCount,11).NumberFormat = "0.00%"
            'Insert TotalStockVolume Formatted
            ws.Cells(TickerCount,12).Value = TotalVolume
            ws.Cells(TickerCount,12).NumberFormat = "0"
            'Check if the current ticker is the Greatest Increase/Decrease and Total and reset if true
            If ws.Cells(GreatestIncrease,11).Value < ws.Cells(TickerCount,11).Value Then
            GreatestIncrease = TickerCount
            End If
            If ws.Cells(GreatestDecrease,11).Value > ws.Cells(TickerCount,11).Value Then
            GreatestDecrease = TickerCount
            End If
            If ws.Cells(GreatestTotal,12).Value < ws.Cells(TickerCount,12).Value Then
            GreatestTotal = TickerCount
            End If

            'Reset variables and Update
            TickerCount = TickerCount + 1
            TotalVolume = 0
            StartingPrice = ws.cells(i+1,3).Value
        End If    
    
    next i
    'Fill In Greatests in Table and Format
    '%Increase
    ws.Cells(2,16).Value = ws.Cells(GreatestIncrease,9).Value
    ws.Cells(2,17).Value = ws.Cells(GreatestIncrease,11).Value
    ws.Cells(2,16).NumberFormat = "0.00%"
    ws.Cells(2,17).NumberFormat = "0.00%"
    '%Decrease
    ws.Cells(3,16).Value = ws.Cells(GreatestDecrease,9).Value
    ws.Cells(3,17).Value = ws.Cells(GreatestDecrease,11).Value
    ws.Cells(3,16).NumberFormat = "0.00%"
    ws.Cells(3,17).NumberFormat = "0.00%"
    'Total
    ws.Cells(4,16).Value = ws.Cells(GreatestTotal,9).Value
    ws.Cells(4,17).Value = ws.Cells(GreatestTotal,12).Value
    ws.Cells(4,16).NumberFormat = "0"
    ws.Cells(4,17).NumberFormat = "0"


next ws

end Sub