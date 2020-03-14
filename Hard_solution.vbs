Sub StockMarketAnalysis()
'This module is a master module which will loop through all the active worksheets
         ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet
         ' Declare lst_Row_Current as last row of the currect active sheet
         Dim lst_Row_Current As Long
         'Declare a variable for holding the current stock ticker
         Dim current_stock_Ticker As String
         'Declare a variable for holding the initial opening price for the stock ticker at the start of the year
         Dim crnt_stock_open_price As Double
         'Initialize this as well
         crnt_stock_open_price = 0
         'Declare a variable for holding the closing price for the stock ticker at the end of the year
         Dim crnt_stock_close_price As Double
         'Initialize this as well
         crnt_stock_close_price = 0
         'Declare a variable for holding the total volume of stock traded over a year
         Dim crnt_stock_total_vol As Double
         'Declare a variable to calculate the yearly change for every stock ticker
         Dim yearly_change As Double
         'Declare a variable to calculate the percentage change for every stock ticker
         Dim prcnt_change As Double
         'Declare variable for biggest percentage increase
         Dim biggest_prcnt_increase As Double
         biggest_prcnt_increase = 0
         Dim biggest_stock_ticker As String
         'Declare variable for biggest percentage decrease
         Dim biggest_prcnt_decrease As Double
         biggest_prcnt_decrease = 0
         Dim least_stock_ticker As String
         'Declare variable for biggest Total Volume
         Dim biggest_total_volume As Double
         biggest_total_volume = 0
         Dim most_traded_stock_ticker As String
         
        ' Keep track of the location for each stock ticker in the summary table
         Dim summary_table_row As Integer
         summary_table_row = 2
         ' Loop through all of the worksheets in the active workbook.
         For Each Current In ActiveWorkbook.Worksheets
            Current.Activate
            'Debug.Print (Current.Name)
            summary_table_row = 2
            ' Set the headers for the summary table as well
             Range("I1").Value = "Ticker"
             Range("J1").Value = "Yearly Change"
             Range("K1").Value = "Percentage Change"
             Range("L1").Value = "Total Stock Volume"
            ' For each Worksheet we need to first find the range of rows we need to work with
            ' Determine the Last Row
              lst_Row_Current = Current.Cells(Rows.Count, 1).End(xlUp).Row
              'Debug.Print (lst_Row_Current)
            ' Now you know the last row and also know that the first row is just the header
            ' so for every sheet you can ignore first row and start from the second row
            ' nested for loop to iterate through all the rows of the current sheet
            For i = 2 To lst_Row_Current
            
                ' Check if we are still within the same stock ticker or not
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'Logic when stock ticker changes. Mainly reset the key variables and add the information to a collection
                    'Set the stock ticker
                    current_stock_Ticker = Cells(i, 1).Value
                    crnt_stock_close_price = Cells(i, 6).Value
                    crnt_stock_total_vol = crnt_stock_total_vol + Cells(i, 7).Value
                        If (crnt_stock_total_vol > biggest_total_volume) Then
                                    biggest_total_volume = crnt_stock_total_vol
                                    most_traded_stock_ticker = current_stock_Ticker
                        End If
                    yearly_change = crnt_stock_close_price - crnt_stock_open_price
                        'If (yearly_change > 0) Then
                           ' If (yearly_change > biggest_prcnt_increase) Then
                                   ' biggest_prcnt_increase = yearly_change
                            'ElseIf (prcnt_change < biggest_prcnt_decrease) Then
                                   ' biggest_prcnt_decrease = prcnt_change
                            'End If
                        'End If
                    'Div by zero error
                        If (crnt_stock_open_price > 0) Then
                            prcnt_change = ((yearly_change) / (crnt_stock_open_price))
                            'Assign it to variables for biggest change positive or negative
                            If (prcnt_change > 0) Then
                                If (prcnt_change > biggest_prcnt_increase) Then
                                    biggest_prcnt_increase = prcnt_change
                                    biggest_stock_ticker = current_stock_Ticker
                                End If
                            ElseIf (prcnt_change < 0) Then
                                    If (prcnt_change < biggest_prcnt_decrease) Then
                                        biggest_prcnt_decrease = prcnt_change
                                        least_stock_ticker = current_stock_Ticker
                                    End If
                            End If
                            
                        Else
                                prcnt_change = 0
                        End If
                    ' Now write all the information out
                    'Range("I" & summary_table_row).NumberFormat = "@"
                    Range("I" & summary_table_row).Value = current_stock_Ticker
                    
                    Range("J" & summary_table_row).NumberFormat = "0.00"
                    Range("J" & summary_table_row).Value = yearly_change
                    ' make sure that the color is blank for
                    If (yearly_change > 0) Then
                        Cells.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    ElseIf (yearly_change < 0) Then
                        Cells.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                    Range("K" & summary_table_row).NumberFormat = "0.00%"
                    Range("K" & summary_table_row).Value = prcnt_change
                    
                    Range("L" & summary_table_row).NumberFormat = "0"
                    Range("L" & summary_table_row).Value = crnt_stock_total_vol
                    
                    'reset variables
                    prcnt_change = 0
                    crnt_stock_total_vol = 0
                    crnt_stock_open_price = 0
                    crnt_stock_close_price = 0
                    yearly_change = 0
                    'incerment row counter
                    summary_table_row = summary_table_row + 1
                    
                Else
                'Assign the open price only if it is zero
                        If (crnt_stock_open_price = 0) Then
                            crnt_stock_open_price = Cells(i, 3).Value
                            'MsgBox (crnt_stock_open_price)
                        End If
                    crnt_stock_total_vol = crnt_stock_total_vol + Cells(i, 7).Value
                End If
                
            Next i
         Current.Range("P1").Value = "Ticker"
         Current.Range("Q1").Value = "Value"
         
         Current.Range("O2").Value = "Greatest % Increase"
         Current.Range("P2").Value = biggest_stock_ticker
         Current.Range("Q2").Value = biggest_prcnt_increase
         Current.Range("Q2").NumberFormat = "0.00%"
         
         Current.Range("O3").Value = "Greatest % Decrease"
         Current.Range("P3").Value = least_stock_ticker
         Current.Range("Q3").Value = biggest_prcnt_decrease
         Current.Range("Q3").NumberFormat = "0.00%"
         
         Current.Range("O4").Value = "Greatest Total Volume"
         Current.Range("P4").Value = most_traded_stock_ticker
         Current.Range("Q4").Value = biggest_total_volume
         
         ' reset indicators
         biggest_stock_ticker = ""
         biggest_prcnt_increase = 0
         least_stock_ticker = ""
         biggest_prcnt_decrease = 0
         most_traded_stock_ticker = ""
         biggest_total_volume = 0
         
         Next Current
         'Hard assignment
'         Worksheets("2016").Range("P1").Value = "Ticker"
'         Worksheets("2016").Range("Q1").Value = "Value"
'
'         Worksheets("2016").Range("O2").Value = "Greatest % Increase"
'         Worksheets("2016").Range("P2").Value = biggest_stock_ticker
'         Worksheets("2016").Range("Q2").Value = biggest_prcnt_increase
'         Worksheets("2016").Range("Q2").NumberFormat = "0.00%"
'
'         Worksheets("2016").Range("O3").Value = "Greatest % Decrease"
'         Worksheets("2016").Range("P3").Value = least_stock_ticker
'         Worksheets("2016").Range("Q3").Value = biggest_prcnt_decrease
'         Worksheets("2016").Range("Q3").NumberFormat = "0.00%"
'
'         Worksheets("2016").Range("O4").Value = "Greatest Total Volume"
'         Worksheets("2016").Range("P4").Value = most_traded_stock_ticker
'         Worksheets("2016").Range("Q4").Value = biggest_total_volume
         
      
End Sub

