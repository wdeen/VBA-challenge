Sub StockAnalysis()
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------
    ' LOOP THROUGH EVERY WORKSHEET IN THE EXCEL WORKBOOK FILE
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------
    For Each ws In Worksheets
    
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' DECLARE VARIABLES TO STORE/READ VALUES FROM WORKSHEET
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
    
        'Relevant variables from the dataset to store & read
        Dim stock_ticker As String                          'Ticker Symbol
        Dim stock_openyear As Double                        'Opening Stock Value of a Ticker at the Beginning of Year
        Dim stock_closeyear As Double                       'Closing Stock Value of a Ticker at the End of Year
        Dim stock_totalvol As Double                        'Total Stock Volume of a Ticker; Double variable type stores numeric values larger than what Long variable type can store
        
        'Variables for ***BONUS*** Challenge
        Dim greatest_percentup_value As Double             'Greatest Percentage Increase (%) in Summary Table
        Dim greatest_percentdown_value As Double           'Greatest Percentage Decrease (%) in Summary Table
        Dim greatest_totalvol_value As Single              'Greatest Total Volume in Summary Table
        
        Dim greatest_percentup_ticker As String            'Ticker Name of Greatest Percentage Increase (%) in Summary Table
        Dim greatest_percentdown_ticker As String          'Ticker Name of Greatest Percentage Decrease (%) in Summary Table
        Dim greatest_totalvol_ticker As String             'Ticker Name of Greatest Total Volume in Summary Table
    
        'Reference Values
        Dim lastrow As Long                                'Number of rows in stock dataset
        Dim summarytable_row As Integer                    'Current row in Summary Table
    
        'Loop counters
        Dim i As Long                                      'For Looping Through Stock Dataset
        Dim j As Long                                      'For Looping Through Summary Table (***BONUS*** Challenge)
    
        'Variables for Conditional Formatting in "Yearly Change" column of Summary Table
        Dim cond_one As FormatCondition
        Dim cond_two As FormatCondition
        Dim cond_three As FormatCondition
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' EXTRACT TOTAL NO. ROWS IN CURRENT WORKSHEET
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Check how many rows in current worksheet
        'MsgBox (lastrow & " Rows in " & ws.Name & " Stock Dataset")
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' SETTING UP AND FORMATTING SUMMARY TABLE
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Set to first row of the Summary Table
        summarytable_row = 1
        
        'Adding Column Names to Header Cells
        ws.Range("I" & summarytable_row).Value = "Ticker"
        ws.Range("J" & summarytable_row).Value = "Yearly Change"
        ws.Range("K" & summarytable_row).Value = "Percent Change"
        ws.Range("L" & summarytable_row).Value = "Total Stock Volume"
        
        'Adding Bold Font to Header Cells in Summary Table
        ws.Range("I" & summarytable_row & ":" & "L" & summarytable_row).Font.Bold = True
        
        'Change column number formats in Summary Table
        ws.Columns("J").NumberFormat = "$0.00"           'Yearly Change Number Format e.g $2.13 (Currency Format)
        ws.Columns("K").NumberFormat = "0.00%"           'Percent Change Number Format e.g 5.67% (Percentage Format)
        ws.Columns("L").NumberFormat = "#,##"            'Total Stock Volume Number Format e.g. 455,675,433 (Numbers with Commas Format)
        
        'Adjusting Column Widths in Summary Table
        ws.Columns("I:L").ColumnWidth = 20
        
        'Horizontally centering all values for all columns in the Summary Table
        ws.Columns("I:L").HorizontalAlignment = xlCenter
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' SETTING UP AND FORMATTING ***BONUS*** TABLE
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Adding Column Names to Header Cells in BONUS Table
        ws.Range("P" & summarytable_row).Value = "Ticker"
        ws.Range("Q" & summarytable_row).Value = "Value"
        
        'Adding Row Names to Row Header Cells in BONUS Table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Adding Bold Font to Column/Row Header Cells in BONUS Table
        ws.Range("P" & summarytable_row & ":" & "Q" & summarytable_row).Font.Bold = True
        ws.Range("O2:O4").Font.Bold = True
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"           'Percent Change Number Format e.g 5.67% (Percentage Format)
        ws.Range("Q4").NumberFormat = "#,##"               'Total Stock Volume Number Format e.g. 455,675,433 (Numbers with Commas Format)
        
        'Adjusting Column Widths in Summary Table
        ws.Columns("O:Q").ColumnWidth = 20
        
        'Horizontally centering all values for all columns in BONUS Table
        ws.Columns("P:Q").HorizontalAlignment = xlCenter
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' KEEP TRACK OF EACH TICKER IN SUMMARY TABLE
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Move on to the 2nd row of the Summary Table
        summarytable_row = summarytable_row + 1
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' LOOP THROUGH DATASET
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Loop through from the 2nd row to the final row of dataset
        For i = 2 To lastrow
        
            'Check if ticker name in current row is the same as ticker name stored in 'stock_ticker'.
            'Important to identify the first row of the current ticker and extract opening stock at beginning of year
            'If not, then...
            If ws.Cells(i, 1).Value <> stock_ticker Then
                stock_ticker = ws.Cells(i, 1).Value         'Store ticker name of current row into string variable
                stock_openyear = ws.Cells(i, 3).Value       'Store Open stock value of current row into double variable
                
            End If
        
            'Triggered when the row after the current does not have the same ticker name.
            'Check if next row does not have the same ticker symbol.  If not, then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stock_closeyear = ws.Cells(i, 6).Value                  'Extract closing stock value in current row at End of Year into double variable
                stock_totalvol = stock_totalvol + ws.Cells(i, 7).Value  'Add current row of stock volume to total stock volume
                
                ws.Range("I" & summarytable_row).Value = stock_ticker                                           'Insert current ticker name into current row in Summary Table
                ws.Range("J" & summarytable_row).Value = stock_closeyear - stock_openyear                       'Difference between closing stock value (End of Year) & opening stock value (Beginning of Year)
                ws.Range("K" & summarytable_row).Value = ((stock_closeyear - stock_openyear) / stock_openyear)  '% Difference between closing stock value (End of Year) & opening stock value (Beginning of Year)
                ws.Range("L" & summarytable_row).Value = stock_totalvol                                         'Insert Total Volume Sales of Current Ticker
                
                'Move on to the next of the Summary Table
                summarytable_row = summarytable_row + 1     'Move to next row in Summary Table for the next ticker
                
                'Reset total stock volume
                stock_totalvol = 0
                
            'Check if next row does not have the same ticker symbol.  If it does, then...
            Else
                stock_totalvol = stock_totalvol + ws.Cells(i, 7).Value                    'Add current row of stock volume to total stock volume
        
            End If
            
        Next i
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' CONDITIONAL FORMATTING SETUP FOR "YEARLY CHANGE" COLUMN
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Conditional Formatting for "Yearly Change" column
        Set cond_one = ws.Range("J2:K" & (summarytable_row - 1)).FormatConditions.Add(xlCellValue, xlLess, "=0")     'If Value is less than 0, use this Conditional Formatting e.g -0.01
        Set cond_two = ws.Range("J2:K" & (summarytable_row - 1)).FormatConditions.Add(xlCellValue, xlGreater, "=0")  'If Value is greater than 0, use this Conditional Formatting e.g. +0.01
        Set cond_three = ws.Range("J2:K" & (summarytable_row - 1)).FormatConditions.Add(xlCellValue, xlEqual, "=0")  'If Value is equal to 0, use this Conditional Formatting e.g. 0.00
       
        'When value in a cell from "Yearly Change" column is < 0
        With cond_one
        
            .Interior.Color = RGB(190, 0, 0)        'Cell colour fill is darker shade of red
            .Font.Color = RGB(255, 255, 255)        'Text font colour is white
        
        End With
        
        
        'When value in a cell from "Yearly Change" column is > 0
        With cond_two
        
            .Interior.Color = RGB(0, 190, 0)       'Cell colour fill is darker shade of green
            .Font.Color = RGB(255, 255, 255)       'Text font colour is white
        
        End With
        
                
        'When value in a cell from "Yearly Change" column is = 0
        With cond_three
        
            .Interior.Color = RGB(100, 100, 100)   'Cell colour fill is darker shade of grey
            .Font.Color = RGB(255, 255, 255)       'Text font colour is white
        
        End With
        
        
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        ' ***BONUS*** CALCULATED VALUES
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Loop through from the 2nd row to the final row of Summary Table
        For j = 2 To summarytable_row
            
            'If % Change value in current row is greater than the value stored in variable, then...
            If ws.Cells(j, 11).Value > greatest_percentup_value Then
                greatest_percentup_value = ws.Cells(j, 11).Value                                    'Store % Change value from current row into variable as Greatest % Increase
                greatest_percentup_ticker = ws.Cells(j, 9).Value                                    'Store ticker name from current row into variable
                
            End If
                
            If ws.Cells(j, 11).Value < greatest_percentdown_value Then
                greatest_percentdown_value = ws.Cells(j, 11).Value                                  'Store % Change value from current row into variable as Greatest % Decrease
                greatest_percentdown_ticker = ws.Cells(j, 9).Value                                  'Store ticker name from current row into variable
                
            End If
            
            If ws.Cells(j, 12).Value > greatest_totalvol_value Then
                greatest_totalvol_value = ws.Cells(j, 12).Value                                     'Store total stock volume from current row into variable as Greatest Total Stock Volume
                greatest_totalvol_ticker = ws.Cells(j, 9).Value                                     'Store ticker name from current row into variable
                
            End If
            
        
        Next j
        
        'Insert stored variables into ***BONUS*** Table
        ws.Range("P2").Value = greatest_percentup_ticker
        ws.Range("P3").Value = greatest_percentdown_ticker
        ws.Range("P4").Value = greatest_totalvol_ticker
        
        ws.Range("Q2").Value = greatest_percentup_value
        ws.Range("Q3").Value = greatest_percentdown_value
        ws.Range("Q4").Value = greatest_totalvol_value
    
        MsgBox ("'" & ws.Name & "' Stock Dataset Analysis Complete")

    Next ws

End Sub
