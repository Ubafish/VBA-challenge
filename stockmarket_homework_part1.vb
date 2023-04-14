Sub stockmarket()
    
    Dim StockName As String
    Dim i As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PriceDiff As Double
    Dim Summary_Table_Row As Integer
    Dim Total_Volume As Double
    
    Summary_Table_Row = 2
    Total_Volume = 0 
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
            StockName = Cells(i, 1).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            ClosePrice = Cells(i, 6).Value
            PriceDiff = ClosePrice - OpenPrice
            PercentChange = PriceDiff / OpenPrice
    
            Range("I" & Summary_Table_Row).Value = StockName
            Range("J" & Summary_Table_Row).Value = PriceDiff
            Range("K" & Summary_Table_Row).Value = FormatPercent(PercentChange)
            Range("L" & Summary_Table_Row).Value = Total_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            OpenPrice = Cells(i + 1, 3).Value
            Total_Volume = 0
            
        ElseIf OpenPrice = 0 Then
            OpenPrice = Cells(i, 3).Value
            
        Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub