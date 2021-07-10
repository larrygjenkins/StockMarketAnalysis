Attribute VB_Name = "Module1"
Sub tickerAnalysis()

For Each ws In Worksheets
    
    'Declaring dimensions of variables
    Dim tickerSymbol As String
    Dim yrOpenPrice As Double
    Dim scndOpenPrice As Double
    Dim yrClosePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim volume As LongLong
    Dim worksheetName As String
    Dim lastRow As Long
    Dim tableRow As Long
    Dim i As Long
    
    worksheetName = ws.Name
    summaryTableRow = 2
    volume = 0
   
    'Adds new column headers to each worksheet
    Worksheets(worksheetName).Range("I1").Value = "Ticker"
    Worksheets(worksheetName).Range("J1").Value = "Yearly Change"
    Worksheets(worksheetName).Range("K1").Value = "Percent Change"
    Worksheets(worksheetName).Range("L1").Value = "Total Stock Volume"
    
    'Changes width of columns to achieve best fit
    Worksheets(worksheetName).Columns("A:L").AutoFit
    
    'Finds last row of data on a worksheet
    lastRow = Worksheets(worksheetName).Cells(Rows.Count, 1).End(xlUp).Row
      
    'Sets value of yrOpenPrice for ticker symbol
    yrOpenPrice = Worksheets(worksheetName).Cells(2, 3).Value
      
    For i = 2 To lastRow
    
        'Checks if consecutive rows show same ticker symbol
        If Worksheets(worksheetName).Cells(i + 1, 1).Value <> Worksheets(worksheetName).Cells(i, 1).Value Then
        
            'sets value of Ticker Symbol variable
            tickerSymbol = Worksheets(worksheetName).Cells(i, 1).Value
            
            'prints the ticker symbol in the summary area of the table
            Worksheets(worksheetName).Range("I" & summaryTableRow).Value = tickerSymbol
            
            'Finds total stock volume per ticker symbol
            volume = volume + Worksheets(worksheetName).Cells(i, 7).Value
            
            'prints the total volume per ticker symbol in the summary area of the table
            Worksheets(worksheetName).Range("L" & summaryTableRow).Value = volume
                                   
            'Sets value of yrClosePrice for ticker symbol
            yrClosePrice = Worksheets(worksheetName).Cells(i, 6).Value
                   
            'Calculates yearly change of price and sets the value of the variable
            yearlyChange = yrClosePrice - yrOpenPrice
            
            'prints yearly change value per ticker symbol in the summary area of the table
            Worksheets(worksheetName).Range("J" & summaryTableRow).Value = yearlyChange
            
                'Checks if yrOpenPrice is 0, which will result in an Overflow error because the calculation will try to divide by 0
                If yrOpenPrice = 0 Then
                    percentChange = 0
                    
                Else
                
                    'If yrOpenPrice is not 0, calculates yearly percent change of price and sets the value of the variable
                    percentChange = Round((((yrClosePrice - yrOpenPrice) / yrOpenPrice) * 100), 2)
                    
                End If
            
            'prints yearly percent change value per ticker symbol in the summary area of the table
            Worksheets(worksheetName).Range("K" & summaryTableRow).Value = percentChange
            
                If Worksheets(worksheetName).Range("J" & summaryTableRow).Value > 0 Then
                    Worksheets(worksheetName).Range("J" & summaryTableRow).Interior.ColorIndex = 4
                    
                ElseIf Worksheets(worksheetName).Range("J" & summaryTableRow).Value < 0 Then
                    Worksheets(worksheetName).Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    
                End If
                      
            'Add one to the summary table row
            summaryTableRow = summaryTableRow + 1
            
            'Reset volume to 0
            volume = 0
            
            'Changes value of yrOpenPrice for ticker symbol
            yrOpenPrice = Worksheets(worksheetName).Cells(i + 1, 3).Value
               
        Else
        volume = volume + Worksheets(worksheetName).Cells(i, 7).Value
        
        End If
    
    Next i
    
    
Next ws

End Sub
