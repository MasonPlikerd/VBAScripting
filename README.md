Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerMaxIncrease As String
    Dim tickerMaxDecrease As String
    Dim tickerMaxVolume As String
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set the last row with data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        tickerMaxIncrease = ""
        tickerMaxDecrease = ""
        tickerMaxVolume = ""
        
        ' Set up headers for results
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percentage Change"
        ws.Cells(1, 11).Value = "Total Volume"
        
        ' Loop through each stock in the sheet
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value
            
            ' Calculate the quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If
            
            ' Record the total volume
            totalVolume = volume
            
            ' Output results to the sheet
            ws.Cells(i, 8).Value = ticker
            ws.Cells(i, 9).Value = quarterlyChange
            ws.Cells(i, 10).Value = percentageChange
            ws.Cells(i, 11).Value = totalVolume
            
            ' Track the stock with the greatest percentage increase, decrease, and volume
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
        Next i
        
        ' Output the greatest increase, decrease, and volume stock to the sheet
        ws.Cells(2, 13).Value = "Greatest % Increase: " & tickerMaxIncrease
        ws.Cells(3, 13).Value = "Greatest % Decrease: " & tickerMaxDecrease
        ws.Cells(4, 13).Value = "Greatest Volume: " & tickerMaxVolume
        
        ' Apply conditional formatting for quarterly change
        Dim rng As Range
        Set rng = ws.Range("I2:I" & lastRow)
        
        ' Positive change in green
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = vbGreen
        End With
        
        ' Negative change in red
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = vbRed
        End With
    Next ws
End Sub

#above is the script ran in VBA to accomplish the taskes of the instructions.
This assignment shows the ability of excel and its VBA function. THis is an intro into the coding part of the bootcamp or a stepping stone. "Dim" is used for imports
