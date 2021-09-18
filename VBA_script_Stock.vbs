Sub loopSHeets()
   For Each ws In Worksheets
    ws.Activate
'=================================================================================================
'Create a script that will loop through all the stocks for one year and output:
    'The ticker symbol.
    'Yearly change from opening price to closing price of a given year.
    'The percent change from opening price to the closing price of a given year.
    'The total stock volume of the stock.

'Get each Ticker symbol and print it in Col 9
    Dim tickerSymbol As String
        'get last row
        Dim lastRow As Double
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'use loop to retrieve each Ticker Symbol, Total Stock Volume, Yearly Change, Percent Change
        Range("I1").Value = "Ticker"
        Range("L1").Value = "Total Stock Vol"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Dim Row As Double
        Dim tickerColCount As Double
        Dim OpeningValue As Double
        Dim ClosingValue As Double
        
        Tickercount = 2
        TotStockVol = 0
        For Row = 2 To lastRow
            If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then  'am I on the LAST row of a Ticker
                Cells(Tickercount, 9).Value = Cells(Row, 1).Value
                Cells(Tickercount, 12).Value = TotStockVol + Cells(Row, 7).Value
                ClosingValue = Cells(Row, 6).Value
                Cells(Tickercount, 10).Value = ClosingValue - OpeningValue
                    'rule out any 0 values in OpeningValue as it can't be divided
                    If OpeningValue = 0 Then
                    Cells(Tickercount, 11).Value = 0
                    Else
                    Cells(Tickercount, 11).Value = (ClosingValue - OpeningValue) / OpeningValue
                    End If
                'adding the last row's volume
                Tickercount = Tickercount + 1
                'reset TotStockVol so it does not keep accumulating across different Tickers
                TotStockVol = 0
             
            ElseIf Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then  'am I on the FIRST row of a Ticker
                OpeningValue = Cells(Row, 3).Value
                TotStockVol = TotStockVol + Cells(Row, 7).Value
            Else  'am I on any OTHER row of Ticker
                TotStockVol = TotStockVol + Cells(Row, 7).Value
            End If
        Next Row
    
'You should also have conditional formatting (green - positive change, red - negative change)
    'color Yearly Change
        Dim LastYearlyChangeRow As Double
        LastYearlyChangeRow = Cells(Rows.Count, 10).End(xlUp).Row
        
        For RowColor = 2 To LastYearlyChangeRow
            If Cells(RowColor, 10).Value > 0 Then
            Cells(RowColor, 10).Interior.ColorIndex = 4
            Else
            Cells(RowColor, 10).Interior.ColorIndex = 3
            End If
        Next RowColor
    'number format Percent Change
        Range("K2:K290").NumberFormat = "0.00%"
    
    
'Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVol As Double
    Dim TotStockLastRow As Integer
    
    'get last row
        Dim SummaryLastRow As Double
        SummaryLastRow = Cells(Rows.Count, 10).End(xlUp).Row
        
                         
    'inserting cells titles
       Range("P1").Value = "Ticker"
       Range("Q1").Value = "Value"
       Range("O2").Value = "Greatest % Increase"
       Range("O3").Value = "Greatest % Decrease"
       Range("O4").Value = "Greatest Total Volume"
       
       
    'Find max and min Change and max TotVol
         Dim rngChange, rngTotVol As Range
         Set rngChange = Range("K2:K" & SummaryLastRow)
         Set rngTotVol = Range("L2:l" & SummaryLastRow)
         GreatestIncrease = Application.WorksheetFunction.Max(rngChange)
         GreatestDecrease = Application.WorksheetFunction.Min(rngChange)
         GreatestVol = Application.WorksheetFunction.Max(rngTotVol)
         Range("Q2").Value = GreatestIncrease
         Range("Q3").Value = GreatestDecrease
         Range("Q4").Value = GreatestVol
         
    'use loop to find related Ticker names
         For i = 2 To SummaryLastRow
           If Cells(i, 10).Value = GreatestIncrease Then
           Range("P2").Value = Cells(i, 9).Value
           ElseIf Cells(i, 10).Value = GreatestDecrease Then
           Range("P3").Value = Cells(i, 9).Value
           ElseIf Cells(i, 12).Value = GreatestVol Then
           Range("P4").Value = Cells(i, 9).Value
           End If
         Next i
     'number format Percent Change
        Range("Q2:Q3").NumberFormat = "0.00%"

'===========================================================================================
    Next ws
End Sub

