VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockdata()
 Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim percentageChange As Double
    Dim quarterlyChange As Double
    Dim firstRow As Long
    Dim lastRowForTicker As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerMaxIncrease As String
    Dim tickerMaxDecrease As String
    Dim tickerMaxVolume As String
    
    For Each ws In ThisWorkbook.Sheets
        maxIncrease = -999999
        maxDecrease = 999999
        maxVolume = 0

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        outputRow = 2 ' For the data output, to skip the heading column
        
        ' Defining the columns for the data results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change (%)"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Initialize the first row data
        currentRow = 2
        Do While currentRow <= lastRow
            ticker = ws.Cells(currentRow, 1).Value
            firstRow = currentRow
            totalVolume = 0 ' Need to adjust for quartetly data format

            ' TO find the last row for the current ticker (before the next ticker starts)
            Do While ws.Cells(currentRow, 1).Value = ticker And currentRow <= lastRow
                totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
                currentRow = currentRow + 1
            Loop
            lastRowForTicker = currentRow - 1

            ' For the open price in first row of the stock ticker
            closePrice = ws.Cells(lastRowForTicker, 6).Value
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If

            ' Output results for the tickers
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).Value = totalVolume

            If percentageChange > maxIncrease Then
                maxIncrease = percentageChange
                tickerMaxIncrease = ticker
            End If

            If percentageChange < maxDecrease Then
                maxDecrease = percentageChange
                tickerMaxDecrease = ticker
            End If

            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                tickerMaxVolume = ticker
            End If

            ' Move to the next output row
            outputRow = outputRow + 1
        Loop

        ' Output the tickers with max values on the relevant sheets
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = tickerMaxIncrease
        ws.Cells(2, 16).Value = Format(maxIncrease, "0.00") & "%"

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = tickerMaxDecrease
        ws.Cells(3, 16).Value = Format(maxDecrease, "0.00") & "%"

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = tickerMaxVolume
        ws.Cells(4, 16).Value = maxVolume
    Next ws
End Sub
