Sub StockAnalysis()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Initialize variables for maximum values
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in the current worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables
        summaryRow = 2
        totalVolume = 0
        
        ' Loop through all rows of data
        For i = 2 To lastRow
            ' Check if the current row is the first row for a new ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Save the ticker symbol and opening price
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Add the stock volume to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if the current row is the last row for the current ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Save the closing price and calculate the yearly change and percent change
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
                End If
                
                ' Output the information to the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format the percent change as a percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Find the stock with the greatest percent increase
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                ' Find the stock with the greatest percent decrease
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                ' Find the stock with the greatest total volume
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                ' Reset variables for the next ticker symbol
                totalVolume = 0
                summaryRow = summaryRow + 1
            End If
        Next i
        yearchangerow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For j = 2 To yearchangerow
        
            If ws.Cells(j, 10).Value > 0 Then
               ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
               ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
               ws.Cells(j, 10).Interior.ColorIndex = 0
            End If
       Next j
      
        ws.Cells(1, 14).Value = "Greatest % Increase"
        ws.Cells(2, 14).Value = "Ticker"
        ws.Cells(2, 15).Value = maxPercentIncreaseTicker
        ws.Cells(3, 14).Value = "Percentage Change"
        ws.Cells(3, 15).Value = maxPercentIncrease
        ws.Cells(3, 15).NumberFormat = "0.00%"
        
        ws.Cells(5, 14).Value = "Greatest % Decrease"
        ws.Cells(6, 14).Value = "Ticker"
        ws.Cells(6, 15).Value = maxPercentDecreaseTicker
        ws.Cells(7, 14).Value = "Percentage Change"
        ws.Cells(7, 15).Value = maxPercentDecrease
        ws.Cells(7, 15).NumberFormat = "0.00%"
        
        ws.Cells(9, 14).Value = "Greatest Total Volume"
        ws.Cells(10, 14).Value = "Ticker"
        ws.Cells(10, 15).Value = maxTotalVolumeTicker
        ws.Cells(11, 14).Value = "Total Volume"
        ws.Cells(11, 15).Value = maxTotalVolume
        
        ws.Columns("A:B").AutoFit
    
    MsgBox "Analysis complete. Please check the 'Results' worksheet for the stocks with the greatest % increase, % decrease, and total volume across all years."
    Next ws
End Sub



