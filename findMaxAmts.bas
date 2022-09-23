Attribute VB_Name = "Module3"
Sub findMaxAmts():
    Dim tickerInc, tickerDec, tickerVol As String
    Dim maxIncrease, maxDecrease, maxVol As Double
    Dim i, lrow As Double
    
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    'determine number of rows
    lrow = Range("I" & Rows.Count).End(xlUp).Row
    
    'print column and row headers and autofit
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest total volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    
    
    'iterate through rows
    For i = 2 To lrow
        Debug.Print Cells(i, 9)
        If Cells(i, 11) > maxIncrease Then
            maxIncrease = Cells(i, 11)
            tickerInc = Cells(i, 9)
        End If
        
        If Cells(i, 11) < maxDecrease Then
            maxDecrease = Cells(i, 11)
            tickerDec = Cells(i, 9)
        End If
        
        If Cells(i, 12) > maxVolume Then
            maxVolume = Cells(i, 12)
            tickerVol = Cells(i, 9)
        End If
    Next i
    
    'print values and autofit
    Range("p2").Value = tickerInc
    Range("p3").Value = tickerDec
    Range("p4").Value = tickerVol
    Range("q2").Value = FormatPercent(maxIncrease, 2)
    Range("q3").Value = FormatPercent(maxDecrease, 2)
    Range("q4").Value = maxVolume
    ActiveSheet.Columns("o:q").AutoFit
    
End Sub
