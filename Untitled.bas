Attribute VB_Name = "Module1"
' Create a script that will loop through all the stocks for one year and output the following information:
    ' The ticker symbol. Step 1
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year. Step 3
    'The % change from opening price at the beginning of a given year to the closing price at the end of that year. Step 4
    'The total stock volume of the stock. Step 2
 Sub StockInfo():
    'variable for stock symbol
    Dim stockName As String
    
    
    'variable for volume total
    Dim volumeTotal As LongLong
    volumeTotal = 0
    
    'variable for summary table
    Dim summaryTableRow As Integer
    summaryTableRow = 2
  
  'variable to hold last row of stock ticker
    Dim lastRow As Long
    
    Cells(1, 9).Value = "Stock Symbols"
    
    Cells(1, 10).Value = "Yearly Change"
    
    Cells(1, 11).Value = "Percentage Change"
    
    Cells(1, 12).Value = "Stock Volume"
    
    
    'count for number of rows
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through stock ticker symbols
        For Row = 2 To lastRow
            
            If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
            
            stockName = Range("A" & Row).Value
            'MsgBox ("Stock Symbol:" & Cells(Row, 1).Value)
            
            volumeTotal = volumeTotal + Range("G" & Row).Value
               
                
                Range("I" & summaryTableRow).Value = stockName
                
            
                Range("L" & summaryTableRow).Value = volumeTotal
                
                    summaryTableRow = summaryTableRow + 1
                    
                    volumeTotal = 0
            
            Else ' if in same stock symbol, add to stock name total
                     volumeTotal = volumeTotal + Range("G" & Row).Value
              
        End If

    
    'variable for yearly price change
    Dim yearlyPriceChange As Double
    Dim priceOpen As Double
    Dim priceClose As Double
        
        'For Row = 2 To lastRow
            
            If Cells(Row, 6).Value <> Cells(Row, 3).Value Then
            priceClose = Range("F" & Row).Value
            priceOpen = Range("C" & Row).Value
            
            
            yearlyPriceChange = priceClose - priceOpen 'Range("F" & Row).Value - Range("C" & Row).Value
            Range("J" & Row).Value = yearlyPriceChange
            'summaryTableRow = summaryTableRow + 1
            
            
            
            Else
            If priceClose = priceOpen Then
            Range("J" & Row).Value = 0
           
            
         End If
    
    'variable for yearly percent change
    Dim yearlyPercentChange As Double
    'yearlyPercentChange = 0
         'For Row = 2 To lastRow
         
            If priceOpen = 0 Then
                    'yearlyPercentChange = yearlyPriceChange / priceOpen
                    Range("K" & Row).Value = yearlyPercentChange
                    
                    If Range("K" & Row).Value = 0 Then
                    yearlyPercentChange = 0
                    
                    
               
                'Next Row
                    
                    End If

        
            End If
            
    End If
    
    
            
    
    'variable for yearly percent change
    'Dim yearlyPercentChange As Double
    
    'variable for total stock volume
    
    'Cells(1, 10).Value = "Yearly Change"
    'Range("J" & summaryTable).Value = yearlyPriceChange
    
    'Cells(1, 11).Value = "Percentage Change"
    'Range("K" & summaryTable).Value = yearlyPercentChange
 
    
            
       Next Row
        
 End Sub
 
 


