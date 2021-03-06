Attribute VB_Name = "Module1"
Sub stock_data_advII()
    
    Dim open_p As Double
    Dim lastRow As Long
    
    For Each ws In Worksheets
    
        num_stock = 2
        total_vol = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest total Volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        G_I = 0
        G_D = 0
        G_V = 0
        
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        open_p = ws.Cells(2, 3).Value
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        
        For i = 2 To lastRow
            total_vol = total_vol + ws.Cells(i, 7).Value
            If open_p = 0 And ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                open_p = ws.Cells(i + 1, 3).Value
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Range("I" & num_stock).Value = ws.Cells(i, 1).Value
                ws.Range("L" & num_stock).Value = total_vol
                ws.Range("J" & num_stock).Value = ws.Cells(i, 6).Value - open_p
                If open_p <> 0 Then
                    ws.Range("K" & num_stock).Value = ws.Range("J" & num_stock).Value / open_p
                Else
                    ws.Range("K" & num_stock).Value = 0
                End If
                
                If ws.Range("J" & num_stock).Value > 0 Then
                    ws.Range("J" & num_stock).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & num_stock).Interior.ColorIndex = 3
                End If
                
                total_vol = 0
                num_stock = num_stock + 1
                open_p = ws.Cells(i + 1, 3).Value
            End If
        Next i
    Next ws
    
End Sub
