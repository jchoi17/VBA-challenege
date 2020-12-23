Attribute VB_Name = "Module1"
Sub SummaryTable()
    
    'dim variables
    Dim ticker As String
    Dim op As Double
    Dim closed As Double
    Dim high As Double
    Dim low As Double
    Dim vol As Double
    Dim total_vol As Double
    Dim ws As Worksheet
    
    'start summary table
    Dim Sum_Table_Row As Integer
    Sum_Table_Row = 2
    
    'loop through worksheets
    For Each ws In Worksheets
    
    
    'define last row  & opening variables
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    op = ws.Cells(2, 3).Value
    'reset summary row
    Sum_Table_Row = 2
        
        'loop through tickers
        For i = 2 To lastrow
        
        'Check for unique values
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' define ticker, closed, high, low, vol variables
            ticker = ws.Cells(i, 1).Value
            closed = ws.Cells(i, 6).Value
            high = ws.Cells(i, 4).Value
            low = ws.Cells(i, 5).Value
            vol = ws.Cells(i, 7).Value
        
            'calculate yearly change
            yearly_change = closed - op
            
            'calculate percent change using if statement
            If op = 0 Then
                percent_change = Null
                Else
                percent_change = yearly_change / op
            End If
            
            percent_change = Format(percent_change, "Percent")
            
            'calculate total volume
            total_vol = total_vol + vol
            
            'fill in table headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            'fill in table data
            ws.Range("I" & Sum_Table_Row).Value = ticker
            ws.Range("J" & Sum_Table_Row).Value = yearly_change
            ws.Range("K" & Sum_Table_Row).Value = percent_change
            ws.Range("L" & Sum_Table_Row).Value = total_vol
            
            'Add row to summary table
            Sum_Table_Row = Sum_Table_Row + 1
            
            'reset values for next row
            total_vol = 0
            op = ws.Cells(i + 1, 3).Value
            
        Else
            
            'Add to total for same ticker
            total_vol = total_vol + ws.Cells(i, 7).Value
            
            
        End If
        
        'Conditional Formatting
        If yearly_change > 0 Then
            ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 4
            Else
            ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 3
        End If
        
        'go on to next ticker
        Next i
        
        
            
    'go on to next worksheet
    Next ws

End Sub



