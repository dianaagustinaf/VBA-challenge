Sub Stock():
    For Each ws In Worksheets:

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim openSt, closeSt As Double
        Dim total, lastrow, count As Integer
        Dim ticker, tempTicker As String
            
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        count = 2  'second row / for results
        total = 0
        
        For i = 2 To lastrow:
            
            vol = ws.Cells(i, 7).Value
            total = total + vol
            
            tempTicker = ws.Cells(i, 1).Value
            'if this ticker is different to the previous one
            If tempTicker <> ws.Cells(i - 1, 1).Value Then
                
                'print ticker in results
                ticker = tempTicker
                ws.Cells(count, 9).Value = ticker
                
                'save first open of the year
                openSt = ws.Cells(i, 3).Value      'C
        
            'if this ticker is different to the next one
            ElseIf tempTicker <> ws.Cells(i + 1, 1).Value Then
                
                'save last close of the year
                closeSt = ws.Cells(i, 6).Value     'F
                                        
                'calculate and print yearly change
                yearlyChange = closeSt - openSt
                ws.Cells(count, 10).Value = yearlyChange
                            
                'calculate and print percent change
                percentChange = (closeSt - openSt) / openSt
                percentChange = Round(percentChange, 2)
                ws.Cells(count, 11).Value = percentChange      'K
                            
                'print total stock volume
                ws.Cells(count, 12).Value = total              'L
                
                'total come back to 0
                total = 0
                'next stock
                count = count + 1

            'if this ticker is equal to the next one -> row++
        
            End If
        
        Next i
        
        'Condition format
        'yearlyChange format: color
        'percentChange format: percent
        
        For i = 2 To lastrow:
                        
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4 'green
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3 'red
            End If
            
            ws.Cells(i, 11).Style = "Percent"
            
        Next i
            
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'BONUS'

        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        Dim tickerIncrease, tickerDecrease, tickerVolume As String
        Dim valueIncrease, valueDecrease, valueVolume As LongLong
                
        valueIncrease = 0
        valueDecrease = 0
        valueVolume = -999999

        For i = 2 To lastrow:

            If ws.Cells(i, 11).Value > valueIncrease Then
                valueIncrease = ws.Cells(i, 11).Value
                tickerIncrease = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value < valueDecrease Then
                valueDecrease = ws.Cells(i, 11).Value
                tickerDecrease = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 12).Value > valueVolume Then
                valueVolume = ws.Cells(i, 12).Value
                tickerVolume = ws.Cells(i, 9).Value
            End If

        Next i

        ws.Range("O2").Value = tickerIncrease
        ws.Range("O3").Value = tickerDecrease
        ws.Range("O4").Value = tickerVolume
    
        ws.Range("P2").Style = "Percent"
        ws.Range("P3").Style = "Percent"
        ws.Range("P2").Value = valueIncrease
        ws.Range("P3").Value = valueDecrease
        ws.Range("P4").Value = valueVolume
        

    Next ws
End Sub
