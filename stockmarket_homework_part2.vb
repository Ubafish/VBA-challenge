Sub stockmarket()
    
    Dim StockName As String
    Dim i As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PriceDiff As Double
    Dim PercentChange As Double
    Dim Summary_Table_Row As Integer
    Dim Total_Volume As Double
    
    Dim MaxIncreaseTicker As String
    Dim MaxIncreaseValue As Double
    
    Dim MaxDecreaseTicker As String
    Dim MaxDecreaseValue As Double
    
    Dim MaxVolumeTicker As String
    Dim MaxVolumeValue As Double
    
    Summary_Table_Row = 2
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    OpenPrice = Cells(2, 3).Value
    
    For i = 2 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
            StockName = Cells(i, 1).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            ClosePrice = Cells(i, 6).Value
            PriceDiff = ClosePrice - OpenPrice
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = PriceDiff / OpenPrice
            End If
            
            Range("I" & Summary_Table_Row).Value = StockName
            Range("J" & Summary_Table_Row).Value = PriceDiff
            Range("K" & Summary_Table_Row).Value = FormatPercent(PercentChange)
            Range("L" & Summary_Table_Row).Value = Total_Volume
            
            ' Determine max increase
            If PercentChange > MaxIncreaseValue Then
                MaxIncreaseTicker = StockName
                MaxIncreaseValue = PercentChange
            End If
            
            ' Determine max decrease
            If PercentChange < MaxDecreaseValue Then
                MaxDecreaseTicker = StockName
                MaxDecreaseValue = PercentChange
            End If
            
            ' Determine max volume
            If Total_Volume > MaxVolumeValue Then
                MaxVolumeTicker = StockName
                MaxVolumeValue = Total_Volume
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            OpenPrice = Cells(i + 1, 3).Value
            Total_Volume = 0
            
        Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
        End If
        
    Next i
    
    ' Output max increase
    Range("P2").Value = MaxIncreaseTicker
    Range("Q2").Value = FormatPercent(MaxIncreaseValue)
    
    ' Output max decrease
    Range("P3").Value = MaxDecreaseTicker
    Range("Q3").Value = FormatPercent(MaxDecreaseValue)
    
    ' Output max volume
    Range("P4").Value = MaxVolumeTicker
    Range("Q4").Value = MaxVolumeValue
    
End Sub

