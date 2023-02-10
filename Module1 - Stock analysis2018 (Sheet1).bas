Attribute VB_Name = "Module1"
Sub stock_analysis18():
    Dim lastRow As Long
    Dim totalVolume As LongLong
    Dim openPrice As Double
    Dim closePrice As Double
    Dim ticker As String
    Dim dollarsChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer
    
    Dim biggestGain As Double
    Dim biggestGainTicker As String
    Dim biggestLoss As Double
    Dim biggestLossTicker As String
    Dim mostVolume As Double
    Dim mostVolumeTicker As String
    
    For Each ws In Worksheets
        ws.Activate
        summaryRow = 2
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        openPrice = Cells(2, 3).Value
        totalVolume = 0
        
        biggestGain = 0
        biggestLoss = 0
        mostVolume = 0
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        For currentRow = 2 To lastRow
            totalVolume = totalVolume + Cells(currentRow, 7)
            
            If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then
            
                ticker = Cells(currentRow, 1).Value
                closePrice = Cells(currentRow, 6).Value
                dollarsChange = closePrice - openPrice
                percentChange = dollarsChange / openPrice
                
                Cells(summaryRow, 9).Value = ticker
                Cells(summaryRow, 10).Value = dollarsChange
                Cells(summaryRow, 11).Value = percentChange
                Cells(summaryRow, 12).Value = totalVolume
                
                If dollarsChange >= 0 Then
                    Cells(summaryRow, 10).Interior.ColorIndex = 4
                Else
                    Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If
                
                If percentChange > biggestGain Then
                    biggestGain = percentChange
                    biggestGainTicker = ticker
                End If
                
                If percentChange < biggestLoss Then
                    biggestLoss = percentChange
                    biggestLossTicker = ticker
                End If
                
                If totalVolume > mostVolume Then
                    mostVolume = totalVolume
                    mostVolumeTicker = ticker
                End If
                
                summaryRow = summaryRow + 1
                
                openPrice = Cells(currentRow + 1, 3).Value
                
                totalVolume = 0
            End If
        Next currentRow
        
        Range("K2:K" & summaryRow).Style = "Percent"
        
        Range("P2").Value = biggestGainTicker
        Range("Q2").Value = biggestGain
        Range("P3").Value = biggestLossTicker
        Range("Q3").Value = biggestLoss
        Range("P4").Value = mostVolumeTicker
        Range("Q4").Value = mostVolume
        
    Next ws
    
    
End Sub

