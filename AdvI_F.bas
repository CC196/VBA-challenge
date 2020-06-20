Attribute VB_Name = "Module111"
Sub stock_data_advI()
    
    Dim open_p As Double
    Dim lastRow As Long
    
    num_stock = 2
    total_vol = 0
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest total Volume"
    
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    G_I = 0
    G_D = 0
    G_V = 0
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    open_p = Cells(2, 3).Value
    Range("K2:K" & lastRow).NumberFormat = "0.00%"
    
    For i = 2 To lastRow
        total_vol = total_vol + Cells(i, 7).Value
        If open_p = 0 And Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            open_p = Cells(i + 1, 3).Value
        End If
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Range("I" & num_stock).Value = Cells(i, 1).Value
            Range("L" & num_stock).Value = total_vol
            Range("J" & num_stock).Value = Cells(i, 6).Value - open_p
            If open_p <> 0 Then
                Percent_c = Range("J" & num_stock).Value / open_p
                Range("K" & num_stock).Value = Percent_c
            Else
                Range("K" & num_stock).Value = 0
            End If
            
            If Range("J" & num_stock).Value > 0 Then
                Range("J" & num_stock).Interior.ColorIndex = 4
            Else
                Range("J" & num_stock).Interior.ColorIndex = 3
            End If
            
            If Percent_c > G_I Then
                G_I = Percent_c
                G_I_name = Range("I" & num_stock).Value
            ElseIf Percent_c < G_D Then
                G_D = Percent_c
                G_D_name = Range("I" & num_stock).Value
            ElseIf total_vol > G_V Then
                G_V = total_vol
                G_V_name = Range("I" & num_stock).Value
            End If
            
            total_vol = 0
            num_stock = num_stock + 1
            open_p = Cells(i + 1, 3).Value
        End If
    Next i
    
    Range("O2").Value = G_I_name
    Range("P2").Value = G_I
    Range("P2").NumberFormat = "0.00%"
    Range("O3").Value = G_D_name
    Range("P3").Value = G_D
    Range("P3").NumberFormat = "0.00%"
    Range("O4").Value = G_V_name
    Range("P4").Value = G_V
    
End Sub
