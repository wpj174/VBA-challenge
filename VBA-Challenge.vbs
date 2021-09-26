Attribute VB_Name = "Module1"
Option Explicit

Sub traverseStockSheets()
    ' Create a script that will loop through all the stocks for one year and output the following information:
    '    Ticker symbol
    '    Yearly change beginning year opening price to ending year closing price
    '    Percent change in yearly price
    '    Total stock volume

    ' Declare required variables
    Dim ws As Worksheet
    Dim numRows, i As Long
    Dim startPrice, endPrice, pctChange, maxIncVal, maxDecVal As Double
    Dim totalVolume, maxVolVal As LongLong
    Dim summaryRow As Integer
    Dim maxIncSym, maxDecSym, maxVolSym As String
    
    For Each ws In Worksheets
    
        'Set up summary table headers for this worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Year Price Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
        'Set start position of summary data
        summaryRow = 2
        
        'Get number of rows in this worksheet
        numRows = ws.Cells(1, 1).End(xlDown).Row
        
        'Set totalVolume accumulator to zero
        totalVolume = 0
        
        'Set starting price of first ticker symbol in this sheet
        startPrice = ws.Cells(2, 3).Value
    
        For i = 2 To numRows    ' Loop through all data rows in this worksheet
            
            'Accumulate total volume each time
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'Check to see if this is the last entry for this ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'Write ticker info to summary table
                
                'Ticker symbol
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value
                
                'Get ending price and calculate price change and percent change
                endPrice = ws.Cells(i, 6).Value
                ws.Cells(summaryRow, 10).Value = (endPrice - startPrice)
                ws.Cells(summaryRow, 10).NumberFormat = "0.00"
                'Format ending price cells
                If ws.Cells(summaryRow, 10).Value < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen
                End If
                
                'check for $0 ending price - trap potential division by 0 errors
                If startPrice = 0 Then
                    ws.Cells(summaryRow, 11).Value = 0
                Else
                    ws.Cells(summaryRow, 11).Value = FormatPercent((endPrice / startPrice) - 1)
                End If
                
                'Write total volume
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                'Increment summary row for next symbol
                summaryRow = summaryRow + 1
                
                'Get new starting price for next symbol
                startPrice = ws.Cells(i + 1, 3).Value
                
                'Reset total volume accumulator
                totalVolume = 0
                
            End If
            
        Next i
        
        'Adjust summary table columns to fit width
        For i = 9 To 12
            ws.Cells(1, i).EntireColumn.AutoFit
        Next i
        
        ' Done compiling summary table
        
        'Scan summary data for greatest % increase, decrease and volume
        ' - this could have been done in the previous loop, but I chose to
        '   show the bonus portion separately.  In addition, the code is more
        '   readable this way.
        
        'Get number of symbols
        numRows = ws.Cells(1, 9).End(xlDown).Row
        
        'Set initial min/max values to zero
        maxIncVal = 0
        maxDecVal = 0
        maxVolVal = 0
        
        'Loop through the symbol summary data on this worksheet
        For i = 2 To numRows
        
            'Check for new max % increase in price
            If ws.Cells(i, 11).Value > maxIncVal Then
                maxIncVal = ws.Cells(i, 11).Value
                maxIncSym = ws.Cells(i, 9).Value
            End If
            
            'Check for new max % decrease in price
            If ws.Cells(i, 11).Value < maxDecVal Then
                maxDecVal = ws.Cells(i, 11).Value
                maxDecSym = ws.Cells(i, 9).Value
            End If
            
            'Check for new max total volume
            If ws.Cells(i, 12).Value > maxVolVal Then
                maxVolVal = ws.Cells(i, 12).Value
                maxVolSym = ws.Cells(i, 9).Value
            End If
        Next i
        
        'Build min/max table
        
        'Headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Max % increase
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncSym
        ws.Cells(2, 16).Value = FormatPercent(maxIncVal)
        
        'Max % decrease
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecSym
        ws.Cells(3, 16).Value = FormatPercent(maxDecVal)
        
        'Max total volume
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = maxVolSym
        ws.Cells(4, 16).Value = maxVolVal
        
        'Adjust column widths to fit
        ws.Cells(1, 14).EntireColumn.AutoFit
        ws.Cells(1, 15).EntireColumn.AutoFit
        ws.Cells(1, 16).EntireColumn.AutoFit
        
        MsgBox ("Finished with sheet " & ws.Name)
        
    Next ws
    
    MsgBox "Done"

    
End Sub

