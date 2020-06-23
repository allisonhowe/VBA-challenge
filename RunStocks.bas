Attribute VB_Name = "Module1"

Sub RunStocks()

    
    Dim last_row As Double
    Dim table_row As Integer
        table_row = 2
    Dim ticker_count As Double
    Dim annual_open As Double
    Dim annual_close As Double
    Dim annual_change As Double
    Dim volume As Double
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Cells(table_row, 9).Value = Cells(i, 1).Value
            
            annual_close = Cells(i, 6).Value
            annual_open = Cells(i - ticker_count, 3).Value
            annual_change = annual_close - annual_open
            
            Cells(table_row, 10).Value = annual_change
            Cells(table_row, 10).NumberFormat = "$0.00"
            If Cells(table_row, 10).Value > 0 Then
                Cells(table_row, 10).Interior.ColorIndex = 10
            ElseIf Cells(table_row, 10).Value < 0 Then
                Cells(table_row, 10).Interior.ColorIndex = 3
            End If
            ' All cells = 0 will not change colors
            
            If annual_open <> 0 Then
                Cells(table_row, 11).Value = annual_change / annual_open
            Else
                Cells(table_row, 11).Value = "n/a"
            End If
            Cells(table_row, 11).NumberFormat = "0.00%"
            
            volume = volume + Cells(i, 7).Value
            Cells(table_row, 12).Value = volume
            Cells(table_row, 12).NumberFormat = "$0,000"
            
            ticker_count = 0
            annual_close = 0
            annual_open = 0
            annual_change = 0
            volume = 0
            table_row = table_row + 1
        
        Else
            ticker_count = ticker_count + 1
            volume = volume + Cells(i, 7).Value
            
        End If
        
    Next i
    
    
    
    last_row_summary = Cells(Rows.Count, 9).End(xlUp).Row
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Dim max As Double
    Dim min As Double
    Dim largest_volume As Double
    
    max = Application.WorksheetFunction.max(Range("K:K"))
    min = Application.WorksheetFunction.min(Range("K:K"))
    largest_volume = Application.WorksheetFunction.max(Range("L:L"))
    
    
    For x = 2 To last_row_summary
    
        If Cells(x, 11).Value = max Then
            Cells(2, 16).Value = Cells(x, 9).Value
            Cells(2, 17).Value = Cells(x, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
            
        ElseIf Cells(x, 11).Value = min Then
            Cells(3, 16).Value = Cells(x, 9).Value
            Cells(3, 17).Value = Cells(x, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
            
        ElseIf Cells(x, 12).Value = largest_volume Then
            Cells(4, 16).Value = Cells(x, 9).Value
            Cells(4, 17).Value = Cells(x, 12).Value
            
        End If
        
    
    Next x



End Sub

