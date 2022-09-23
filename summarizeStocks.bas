Attribute VB_Name = "Module1"
Sub summarizeStocks():
    'declare variables
    Dim i, lrow, sumRow, currDate, minDate, maxDate As Long
    Dim ticker As String
    Dim openingPrice, closingPrice, change, pctChange, stockVol As Double
    
    'assign values to variables
    minDate = 99999999
    maxDate = 0
    stockVol = 0#
    sumRow = 2
    
    'determine number of rows
    lrow = Range("A" & Rows.Count).End(xlUp).Row
    
    'print headers
    [i1].Value = "Ticker"
    [j1].Value = "Yearly Change"
    [k1].Value = "Percent Change"
    [l1].Value = "Total Stock Volume"
    ActiveSheet.Columns("i:l").AutoFit
    
    'iterate through rows
    For i = 2 To lrow
        currDate = CLng(Cells(i, 2))
        
        'calculate stock volume
        stockVol = stockVol + Cells(i, 7).Value
        
        'store first recorded date and opening price
            If currDate <= minDate Then
                minDate = currDate
                openingPrice = Cells(i, 3).Value
            End If
        
        'for each unique ticker symbol
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            
            'store last recorded date and closing price
            If currDate >= maxDate Then
                maxDate = currDate
                closingPrice = Cells(i, 6).Value
            End If
            
            'record ticker symbol
            ticker = Cells(i, 1)
            
            'calculate yearly change in price; round to 2 decimals
            change = Round(closingPrice - openingPrice, 2)
    
            'calculate yearly percent change in price
            pctChange = FormatPercent(change / openingPrice, 2)
            
            'print ticker summary info to spreadsheet
            Debug.Print ticker; minDate; maxDate; openingPrice; closingPrice; change; pctChange; stockVol
            Cells(sumRow, 9) = ticker
            Cells(sumRow, 10) = change
            Cells(sumRow, 11) = pctChange
            Cells(sumRow, 12) = stockVol
            
            'apply formatting to yearly change column
            'if change is positive, then highlight cell green, else highlight cell red
            If change >= 0 Then
                Cells(sumRow, 10).Interior.ColorIndex = 4
            Else
                Cells(sumRow, 10).Interior.ColorIndex = 3
            End If
            'set number of decimals to two in yearly change column
            Cells(sumRow, 10).NumberFormat = "0.00"
            
            'reset variables
            minDate = 99999999
            maxDate = 0
            openingPrice = 0
            closingPrice = 0
            stockVol = 0
            
            'increment summary table row number
            sumRow = sumRow + 1
            
            'Debug.Print ticker; minDate; maxDate; openingPrice; closingPrice; change; stockVol
        End If
    Next i
End Sub

