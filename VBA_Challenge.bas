Attribute VB_Name = "Module1"
Sub ticker_table()
    
    'Declare variables:
    Dim sheet_number As Integer
    Dim table_number As Integer
    Dim row_number As Long
    Dim ticker As String
    Dim opening_value As Double
    Dim closing_value As Double
    Dim volume_sum As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    sheet_number = 1

    'While loop to cycle through all sheets
    Do While sheet_number < Sheets.Count + 1
        
        'Go to current sheet
        Sheets(sheet_number).Select
        
        'Set up table on current sheet
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        table_number = 2
        row_number = 2
        
        'While loop to cycle through tickers until there are none left
        Do While Cells(row_number, 1).Value <> ""
            'Store opening value
            opening_value = Cells(row_number, 3).Value
            'Reset volume sum
            volume_sum = 0
            
            'While loop to cycle through rows until Ticker changes
            Do While Cells(row_number, 1).Value = Cells(row_number + 1, 1).Value
                'Cumulative sum of volume
                volume_sum = volume_sum + Cells(row_number, 7).Value
                'Move to next row
                row_number = row_number + 1
                
            Loop
                
            'Store ticker value
            ticker = Cells(row_number, 1).Value
            'Add last volume to sum
            volume_sum = volume_sum + Cells(row_number, 7).Value
            'Store closing value
            closing_value = Cells(row_number, 6).Value
            'Calculate yearly change
            yearly_change = closing_value - opening_value
            'Calculate percent change
            percent_change = yearly_change / opening_value

            'Print ticker data to table
            Cells(table_number, 9).Value = ticker
            Cells(table_number, 10).Value = yearly_change
            Cells(table_number, 11).Value = percent_change
            Cells(table_number, 12).Value = volume_sum
            'Fromat yearly cahnge
            If yearly_change >= 0 Then
                Cells(table_number, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(table_number, 10).Interior.ColorIndex = 3
            End If
            table_number = table_number + 1
            'Move to first row in next ticker
            row_number = row_number + 1

            
        Loop
        
        'Create table for extreme values
        Range("N1").Value = "Extremes"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("P2").Value = "=MAX(K:K)"
        Range("P3").Value = "=MIN(K:K)"
        Range("P4").Value = "=MAX(L:L)"
        Range("O2").Value = "=XLOOKUP(P2,K:K,I:I)"
        Range("O3").Value = "=XLOOKUP(P3,K:K,I:I)"
        Range("O4").Value = "=XLOOKUP(P4,L:L,I:I)"
        
        'Formatting for current sheet
        'Format Percentages
        Columns("K:K").NumberFormat = "0.00%"
        Range("P2:P3").NumberFormat = "0.00%"
        
        'Auto fit columns
        Range("J:J").EntireColumn.AutoFit
        Range("K:K").EntireColumn.AutoFit
        Range("L:L").EntireColumn.AutoFit
        Range("M:M").EntireColumn.AutoFit
        Range("N:N").EntireColumn.AutoFit
        Range("P:P").EntireColumn.AutoFit

        'Select A1
        Cells(1, 1).Select
        
        'Next sheet
        sheet_number = sheet_number + 1
        
    Loop
    
    'Go to A1 on sheet 1
    Sheet1.Activate
    Cells(9, 1).Select
    
End Sub
