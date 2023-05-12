Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim last_row As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim total_volume As Double
    Dim summary_table_row As Long
    Dim max_total_volume_ticker As String
    Dim max_total_volume As Double
    Dim max_percent_increase_ticker As String
    Dim max_percent_increase As Double
    Dim max_percent_decrease_ticker As String
    Dim max_percent_decrease As Double
    
    
    For Each ws In ActiveWorkbook.Worksheets
        
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

        ticker = ""
        yearly_change = 0
        percent_change = 0
        open_price = 0
        close_price = 0
        total_volume = 0
        summary_table_row = 2
        
        last_row = ws.Range("A" & ws.Rows.Count).End(xlUp).Row

        
        For i = 2 To last_row
            
            
            If ws.Cells(i, 1).Value <> ticker Then
                
                
                ticker = ws.Cells(i, 1).Value
                
                
                open_price = ws.Cells(i, 3).Value
                
            End If
            
            
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                
                close_price = ws.Cells(i, 6).Value
                
                
                yearly_change = close_price - open_price
                
               
                percent_change = yearly_change / open_price
                
                
            If yearly_change < 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            
            ElseIf yearly_change > 0 Then
            
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            End If
                
                ws.Cells(summary_table_row, 9).Value = ticker
                ws.Cells(summary_table_row, 10).Value = yearly_change
                ws.Cells(summary_table_row, 11).Value = percent_change
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                ws.Cells(summary_table_row, 12).Value = total_volume
                
                If total_volume > max_total_volume Then
                    max_total_volume_ticker = ticker
                    max_total_volume = total_volume
                End If
                
                If percent_change > max_percent_increase Then
                    max_percent_increase_ticker = ticker
                    max_percent_increase = percent_change
                End If
                
                If percent_change < max_percent_decrease Then
                    max_percent_decrease_ticker = ticker
                    max_percent_decrease = percent_change
                End If
                
                ticker = ""
                yearly_change = 0
                percent_change = 0
                open_price = 0
                close_price = 0
                total_volume = 0
                summary_table_row = summary_table_row + 1
                
            End If
            
        Next i

        ws.Columns("I:L").AutoFit
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(2, 15).Value = max_percent_increase_ticker
        ws.Cells(2, 16).Value = max_percent_increase
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = max_percent_decrease_ticker
        ws.Cells(3, 16).Value = max_percent_decrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = max_total_volume_ticker
        ws.Cells(4, 16).Value = max_total_volume

    Next ws

End Sub
