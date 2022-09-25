Attribute VB_Name = "Module2"
Sub Ticker_Summary()
        '----Loop through workbook sheets----------

        Dim ws As Worksheet
        Dim starting_ws As Worksheet
        Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
        For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        '----Begin of Code for each sheet-------
        Dim i As Long 'best practice

        'Hardcoding last row value for first column
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        Dim TotalVolume
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim CurrentTik As String

        'Used to add rows under current row where values will be printed
        'This will allow for a row of data to be printed for each <ticker> value
        Dim Result_Row As Long

        'variables for bonus
        Dim Max_Perc_Change  As Double
        Dim Max_Perc_Stock As String

        Dim Min_Perc_Change  As Double
        Dim Min_Perc_Stock As String

        Dim Max_Volume
        Dim Max_Volume_Stock As String

        'Column headers for initial result columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"


        Result_Row = 2
        TotalVolume = 0
        CurrentTik = Cells(2, 1).Value


                '----------This prints data into the new rows and columns to the right of the raw information---------

                'Loop starts at 2 in order to skip column labels
                For i = 2 To LastRow

                        '-----VOLUME SUM IF STATEMENT------ This will loop through and add all the <vol> values for each <ticker>
                        If Cells(i, 1).Value = CurrentTik Then
                                TotalVolume = TotalVolume + Cells(i, 7).Value
                        End If
                
                        ' This IF statement will return the first <open> value for each individual <ticker> value
                        If Cells(i - 1, 1).Value <> CurrentTik Then
                                OpenPrice = Cells(i, 3).Value
                        End If

                        'This detects a change in the value for <ticker> in the next row
                        If Cells(i + 1, 1).Value <> CurrentTik Then
                                'This populates the "Ticker" column with the current <ticker> value
                                Cells(Result_Row, 9).Value = CurrentTik
                                'This populates the "Total Volume" column with the sum of the <vol> column for the current<ticker>
                                'This value is created by the VOLUME SUM IF STATEMENT
                                Cells(Result_Row, 12).Value = TotalVolume
                                'Last <close> value for the current <ticker>
                                ClosePrice = Cells(i, 6).Value
                                'Calculates "Yearly Change" column for current <ticker>
                                Cells(Result_Row, 10).Value = ClosePrice - OpenPrice
                                'Calculates "Percent Change" for current <ticker>
                                Cells(Result_Row, 11).Value = (ClosePrice - OpenPrice) / OpenPrice
                                'Reset TotalVolume to zero
                                TotalVolume = 0
                                'Changes the CurrentTik to the next <ticker> value
                                CurrentTik = Cells(i + 1, 1).Value
                                'Sets the Result_Row to the next row in the summary table so the next <ticker> information will print there.
                                Result_Row = Result_Row + 1
                                
                        End If
                Next i

        'Formatting for first summary table. The LastColorRow function will now apply to row "J" for the rest of the subroutine
        'This will allow for proper conditional formatting in the first summary table and the creation of the secondary summary table
        LastColorRow = Cells(Rows.Count, 10).End(xlUp).Row
        Max_Perc_Change = 0
        Min_Perc_Change = 0
        Max_Volume = 0

                'New loop "c" for Row "J". Again using a value of "2" to skip column headers
                For c = 2 To LastColorRow

                        'Conditionals for Color Fill
                        If Cells(c, 10).Value >= 0 Then
                                Cells(c, 10).Interior.ColorIndex = 4
                                ElseIf Cells(c, 10).Value < 0 Then
                                Cells(c, 10).Interior.ColorIndex = 3
                        End If
                        'Bonus Exercise (Creation or secondary summary table)
                        'Columns for bonus exercise. Loop c is still used because these values will be derived from looping through
                        'the first summary table where the 'LastColorRow' variable returns the proper value for the last populated row
                        If Cells(c, 11).Value > Max_Perc_Change Then
                                Max_Perc_Change = Cells(c, 11).Value
                                Max_Perc_Stock = Cells(c, 9).Value
                        End If
                                
                        If Cells(c, 11).Value < Min_Perc_Change Then
                                Min_Perc_Change = Cells(c, 11).Value
                                Min_Perc_Stock = Cells(c, 9).Value
                        End If
                                
                        If Cells(c, 12).Value > Max_Volume Then
                                Max_Volume = Cells(c, 12).Value
                                Max_Volume_Stock = Cells(c, 9).Value
                        End If


                'End of second loop "c"
                Next c

                        'Hardcoded Row and Column headings for secondary summary table
                        Cells(2, 15).Value = "Greatest % Increase"
                        Cells(3, 15).Value = "Greatest % Decrease"
                        Cells(4, 15).Value = "Greatest Total Volume"
                        Cells(1, 16).Value = "Ticker"
                        Cells(1, 17).Value = "Value"
                        'Ticker string and formatted value for percentage increase for "Greatest % Increase"
                        Cells(2, 16).Value = Max_Perc_Stock
                        Cells(2, 17).Value = Max_Perc_Change
                        Cells(2, 17).NumberFormat = "0.00%"
                                
                        'Ticker string and formatted value for percentage decrease for "Greatest % Decrease"
                        Cells(3, 16).Value = Min_Perc_Stock
                        Cells(3, 17).Value = Min_Perc_Change
                        Cells(3, 17).NumberFormat = "0.00%"

                        'Ticker string and formatted value for "Greatest Total Volume" value
                        Cells(4, 16).Value = Max_Volume_Stock
                        Cells(4, 17).Value = Max_Volume
                        Cells(4, 17).NumberFormat = "0"
                                
                        'Autofit applied to column for readiblity after all data is printed to sheets
                        Columns("A:T").AutoFit
                    
        '-------End of Looping through Workbook Sheets----------------
        Next
        starting_ws.Activate
        '---------------------------

End Sub

