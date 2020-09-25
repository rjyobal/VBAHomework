Sub worksheetLoop()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call createSummary
    Next
    Application.ScreenUpdating = True
End Sub

Sub createSummary()
    'Create Labeling for Summary Table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
   
    'Fill out Summary Table
    Dim ticker As String
    Dim ticker_total As Double
    Dim summ_table_row As Long
    Dim firstValue As Double
    Dim lastrow As Long
    
    Dim lastrowsum As Long
    
    summ_table_row = 1
    summ_table_row2 = 1
    ticker_total = 0
    
    lastrow = (Cells(Rows.Count, "A").End(xlUp).Row)
    'MsgBox (lastrow)
    
    For i = 2 To lastrow
        ticker_total = ticker_total + Cells(i, 7).Value
        '-- IF creates summary table --
        '-Fill outs when there is only one record for Ticker
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Fill outs First Value column
            summ_table_row2 = summ_table_row2 + 1
                'Cells(summ_table_row2, 19).Value = Cells(i, 3).Value 'Validation Column to Comment
            'Fill outs Total Volume column
            ticker = Cells(i, 1).Value
            summ_table_row = summ_table_row + 1
            Cells(summ_table_row, 9).Value = ticker
            Cells(summ_table_row, 12).Value = ticker_total
            ticker_total = 0
            'Get Last Value from Amount
                'Cells(summ_table_row, 20).Value = Cells(i, 6).Value 'Validation Column to Comment
            'Calculate Yearly Change of First Value and Last Value
            Cells(summ_table_row, 10).Value = Cells(i, 6).Value - Cells(i, 3).Value
            'Calculate Percent Change of First Value and Last Value
            Cells(summ_table_row, 11).Value = FormatPercent((Cells(i, 6).Value - firstValue) / firstValue)
            'Color the cell based on value
            If (Cells(i, 6).Value - firstValue) <= 0 Then
                Cells(summ_table_row, 10).Interior.ColorIndex = 3
            Else
                Cells(summ_table_row, 10).Interior.ColorIndex = 4
            End If
            
        '-Fill outs when there are multiple records for Ticker
        'Fill outs Total Volume column by Ticker
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Fill outs Total Volume column
            ticker = Cells(i, 1).Value
            summ_table_row = summ_table_row + 1
            Cells(summ_table_row, 9).Value = ticker
            Cells(summ_table_row, 12).Value = ticker_total
            ticker_total = 0
            'Get Last Value from Close
                'Cells(summ_table_row, 20).Value = Cells(i, 6).Value 'Validation Column to Comment
            'Calculate Yearly Change of First Value and Last Value
            Cells(summ_table_row, 10).Value = Cells(i, 6).Value - firstValue
            'Calculate Percent Change of First Value and Last Value
            
            If Cells(i, 6).Value <> 0 Then
                If firstValue = 0 Then
                    Cells(summ_table_row, 11).Value = FormatPercent(Cells(i, 6).Value / 100)
                Else
                    Cells(summ_table_row, 11).Value = FormatPercent((Cells(i, 6).Value - firstValue) / firstValue)
                End If
            Else
                Cells(summ_table_row, 11).Value = FormatPercent(0)
            End If
            
            'Color the cell based on value
            If (Cells(i, 6).Value - firstValue) <= 0 Then
                Cells(summ_table_row, 10).Interior.ColorIndex = 3
            Else
                Cells(summ_table_row, 10).Interior.ColorIndex = 4
            End If
            
        'Fill outs First Value when there are multiple records for Ticker
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            summ_table_row2 = summ_table_row2 + 1
                'Cells(summ_table_row2, 19).Value = Cells(i, 3).Value 'Validation Column to Comment
            firstValue = Cells(i, 3).Value
        End If
    Next i
    
    'Challenge Coding
    'Create Labeling for Summary Table
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    lastrowsum = (Cells(Rows.Count, "I").End(xlUp).Row)
    
    Dim arrayMax() As Double
    ReDim arrayMax(lastrowsum)
    
    'Fill out Array with Percent Change Values
    For j = 2 To lastrowsum
        arrayMax(j) = Cells(j, 11).Value
    Next j
    'Get Max and Min from Array and Print it
    Range("Q2").Value = FormatPercent(WorksheetFunction.Max(arrayMax))
    maxValue = WorksheetFunction.Max(arrayMax)
    Range("Q3").Value = FormatPercent(WorksheetFunction.Min(arrayMax))
    minValue = WorksheetFunction.Min(arrayMax)
    'Get the Ticker for Max and Min Percentages
    For j = 2 To lastrowsum
        If Cells(j, 11).Value = maxValue Then
            Range("P2").Value = Cells(j, 9).Value
        ElseIf Cells(j, 11).Value = minValue Then
            Range("P3").Value = Cells(j, 9).Value
        End If
    Next j
    
    'Fill out Array with Stock Volume Values
    For k = 2 To lastrowsum
        arrayMax(k) = Cells(k, 12).Value
    Next k
    'Get Max and Min from Array and Print it
    Range("Q4").Value = WorksheetFunction.Max(arrayMax)
    maxValue = WorksheetFunction.Max(arrayMax)
    'Get the Ticker for Max Volume
    For j = 2 To lastrowsum
        If Cells(j, 12).Value = maxValue Then
            Range("P4").Value = Cells(j, 9).Value
        End If
    Next j
    
End Sub




