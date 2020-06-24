Attribute VB_Name = "Module1"
Sub stocks()

Dim i As Long

Dim ticker As String

Dim total_stock_volume As LongLong
total_stock_volume = 0

Dim summary_table_row As Integer
summary_table_row = 2

Dim price_already_captured  As Boolean

Dim close_price As Integer

Dim ws As Worksheet

Dim last_row As Long

Dim current_vol As Integer


For Each ws In Worksheets
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
summary_table_row = 2

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"


    
    For i = 2 To last_row
    
        If price_already_captured = False Then
            Dim open_price As Double
            open_price = ws.Cells(i, 3).Value
    
            price_already_captured = True
        End If
    
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            ticker = ws.Cells(i, 1).Value
            
          
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            price_already_captured = False
            
            yearly_change = ws.Cells(i, 6).Value - open_price
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = yearly_change
            If open_price = 0 Then
                per_change = 0
            Else
                per_change = (yearly_change / open_price) * 100
            End If
            ws.Range("K" & summary_table_row).Value = per_change
            ws.Range("L" & summary_table_row).Value = total_stock_volume
           
            summary_table_row = summary_table_row + 1
            total_stock_volume = 0
        Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        End If
        
        If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
        

    Next i

    
Next ws

End Sub

