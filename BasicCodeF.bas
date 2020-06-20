Attribute VB_Name = "Module11"
Sub stock_data()
    
    Dim open_p As Double
    Dim lastRow As Long
    
    num_stock = 2
    total_vol = 0
    
    'to create column name
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    open_p = Cells(2, 3).Value
    
    For i = 2 To lastRow
        total_vol = total_vol + Cells(i, 7).Value
        'check if opening price is 0
        If open_p = 0 And Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            open_p = Cells(i + 1, 3).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Range("I" & num_stock).Value = Cells(i, 1).Value
            Range("L" & num_stock).Value = total_vol
            Range("J" & num_stock).Value = Cells(i, 6).Value - open_p
            'prevent overflow
            If open_p <> 0 Then
                Range("K" & num_stock).Value = Range("J" & num_stock).Value / open_p
                Range("K" & num_stock).NumberFormat = "0.00%"
            Else
                Range("K" & num_stock).Value = 0
                Range("K" & num_stock).NumberFormat = "0.00%"
            End If
            
            If Range("J" & num_stock).Value > 0 Then
                Range("J" & num_stock).Interior.ColorIndex = 4
            Else
                Range("J" & num_stock).Interior.ColorIndex = 3
            End If
            
            total_vol = 0
            num_stock = num_stock + 1
            open_p = Cells(i + 1, 3).Value

        End If
    Next i
    
End Sub
