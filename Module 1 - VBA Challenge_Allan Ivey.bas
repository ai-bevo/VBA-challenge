Attribute VB_Name = "Module1"
Sub StockChallenge()

'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

'declare variables

Dim ws As Worksheet
Dim starting_ws As Worksheet

Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Dim col, i, Ticker, TickerCount As Integer
    Dim Volume, MaxPercent, MinPercent, MaxVolume As Double
    Dim MaxTicker, MinTicker, MaxVTicker As String
      'conditional formatting variables
    Dim rg As Range
    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
    
    Cells.ClearFormats 'Clear formats
    Range("J:T").ClearContents 'Rather than manually deleting cells each time during testing
    
    Range("C2:F2", Range("C2:F2").End(xlDown)).Style = "Currency" 'Format stock data as currency prices just because it should be
    
    col = 1 'column to analyze, probably don't need this, but oh well
    TickerCount = 2 'starts looking at row 2 for ticker name
    Volume = 0 'sets volume at zero
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'defines lastrow dynamically
    
    Cells(1, 10).Value = "Ticker" 'labels column as Ticker...also could have used range for these.
    Cells(1, 11).Value = "Open Price" 'lables column for Open price
    Cells(1, 12).Value = "Close Price" 'lables column for Close price
    Cells(1, 13).Value = "1 Year - Price Change" 'labels columns for price annual price difference
    Cells(1, 14).Value = "1 Year - Percent Change" 'labels columns for price percentage change
    Cells(1, 15).Value = "Volume" 'labels column as Volume
    
    Cells(TickerCount, 10).Value = Cells(TickerCount, 1) 'enters first ticker name
    Cells(TickerCount, 11).Value = Cells(TickerCount, 3) 'enters first open value
    
    'populate open and close prices in this loop
        For i = 2 To lastrow
        
            If Cells(i + 1, col).Value <> Cells(i, col).Value Then        'comparing ticker cells vertically
            TickerCount = TickerCount + 1                                 'counting for each change in ticker value
            Cells(TickerCount, 10).Value = Cells(i + 1, col)               'Populating each Ticker name as count grows in ticker column
            Cells(TickerCount - 1, 15).Value = Volume + Cells(i, 7).Value 'Adds last volume value
            Volume = 0                                                    'When ticker changes volume is reset to zero
            
            Cells(TickerCount, 11).Value = Cells(i + 1, 3).Value          'Adds opening price from first entry for new ticker
            Cells(TickerCount - 1, 12).Value = Cells(i, 6).Value          'Adds closing price from last entry of new ticker
            
            Else
            
            Volume = Volume + Cells(i, 7).Value 'When ticker doesn't change volume value is being summed
            
            End If
        Next i
    
    'perform math with the populated prices in this loop
    lastrow_2 = Cells(Rows.Count, 11).End(xlUp).Row
    
        For i = 2 To lastrow_2
    
            Cells(i, 13).Value = (Cells(i, 12).Value - Cells(i, 11).Value) 'calculates change YoY between close and open price
            Cells(i, 14).Value = (Cells(i, 13).Value / Cells(i, 11).Value) 'calculates % change YoY between close and open price
                
        Next i
    'Add lables for greatest increase, decrease, and greatest volume
    Range("R2").Value = "Greatest % Increase"
    Range("R3").Value = "Greatest % Decrease"
    Range("R4").Value = "Greatest Total Volume"
    Range("S1").Value = "Ticker"
    Range("T1").Value = "Value"
    
    'Find the max, min percents, and the largest volume
    MaxPercent = Application.WorksheetFunction.Max(Range("N2", Range("N2").End(xlDown)))
    MinPercent = Application.WorksheetFunction.Min(Range("N2", Range("N2").End(xlDown)))
    MaxVolume = Application.WorksheetFunction.Max(Range("O2", Range("O2").End(xlDown)))
    
    'populate values to the worksheet
    Range("T2").Value = MaxPercent
    Range("T3").Value = MinPercent
    Range("T4").Value = MaxVolume
    
    'Almost went crazy trying to use Vlookup for way too long
    'then finally found out Vlookup won't look to the left. Figured out how to use Index & Match to return the corresponding ticker for the max, min, and maxvolume
    MaxTicker = Application.WorksheetFunction.Index(Range("J:J"), Application.WorksheetFunction.Match(MaxPercent, Range("N:N"), 0))
    MinTicker = Application.WorksheetFunction.Index(Range("J:J"), Application.WorksheetFunction.Match(MinPercent, Range("N:N"), 0))
    MaxVTicker = Application.WorksheetFunction.Index(Range("J:J"), Application.WorksheetFunction.Match(MaxVolume, Range("O:O"), 0))
    
    'Populates cells with values that align with the index column and the row that matches the max and min values from the calculated data
    Range("S2").Value = MaxTicker
    Range("S3").Value = MinTicker
    Range("S4").Value = MaxVTicker
    
    
    'setting the range as a variable just cleans up the following commands for conditional formatting
    Set rg = Range("M2", Range("N2").End(xlDown))
    
    rg.FormatConditions.Delete 'clears any previous conditional formatting
    
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)
    Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, 0)
    
    'cond1 format - "with" allows me to apply multiple formats at the same time
    With cond1
    .Interior.Color = vbGreen
    .Font.Color = vbBlack
    End With
    
    'cond2 format
    With cond2
    .Interior.Color = vbRed
    .Font.Color = vbBlack
    End With
    
    'cond3 format
    With cond3
    .Interior.Color = vbYellow
    .Font.Color = vbBlack
    End With
    
    
    'format populated ranges
    Range("K2:L2", Range("K2:L2").End(xlDown)).Style = "Currency" 'ranges for currency
    Range("N2", Range("N2").End(xlDown)).NumberFormat = "0.00%" 'ranges for percentages
    Range("T2:T3").NumberFormat = "0.00%" 'range for percentages
    Range("J1:T1").Font.Bold = True 'make header row bold text
  
    Range("J1:T1", Range("J1:T1").End(xlDown)).Columns.AutoFit 'autofit columns to finish
    

Next

End Sub

