Sub VBA_challenge()

' MsgBox " Congratulations - Start "

' INSTRUCTIONS:
    ' Create a script that will loop through all the stocks for one year for each run and take the following information.
        ' The ticker symbol.
        ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The total stock volume of the stock.
    ' You should also have conditional formatting that will highlight positive change in green and negative change in red.

' CHALLENGES:
    ' 1.  The solution will also be able to return the stock with the
    '           "Greatest % increase"
    '           "Greatest % Decrease"
    '           "Greatest total volume"
    ' 2.  Make the appropriate adjustments to the VBA script that will allow it to run on every worksheet,
    '       i.e., every year, just by running the VBA script once.

' ============================================================================================================================
' ||    Declare variables                                                                                                   ||
' ============================================================================================================================
Dim Ticker As String                ' Set variable for the ticker symbol
Dim Open_Price As Double            ' Set variable for the opening stock price
Open_Price = 0

Dim Close_Price As Double           ' Set variable for the closing stock price
Close_Price = 0

Dim Volume As Double                ' Set variable for the stock trading volumn and initial number traded for counter
Volume = 0

Dim Summary_Table_Row As Integer    ' Keep track of the location for each Ticker Symbol in the Summary Table
Dim Last_Row As Long                ' Set variable to find last row of Ticker data
Dim Last_Row_Summary_Table As Long  ' Set variable to find last row of Summary Table data
Dim Yearly_Change As Double         ' Set variable for the Yearly Change number in Summary Table
Dim Percent_Change As Double        ' Set variable for the Percent Change number in Summary Table
Dim Previous_Amount As Double
Previuos_Amount = 2

Dim ws As Worksheet                 ' Set variable for Worksheet

Dim Greatest_Percent_Increase_Value As Double       ' Set variable for Greatest % Increase Value
Dim Greatest_Percent_Decrease_Value As Double       ' Set variable for Greatest % Decrease Value
Dim Greatest_Total_Volumn_Ticker As String          ' Set variable for Greatest Total Volume Ticker
Dim Greatest_Total_Volumn_Value As Double           ' Set variable for Greatest Total Volume Value

' Set Boolean to test if Stock Ticker Symbol repeats on next row
 Dim Flag As Boolean
 Flag = False

' Column headings for new data in Summary Table - columns I:N
    ' Column I - Ticker
    ' Column J - Open Price
    ' Column K - Close Price
    ' Column L - Yearly Change
    ' Column M - Percent Change
    ' Column N - Total Stock Volumn

' Column headings for new data in Greatest Stocks Table - columns P:R
    ' Column P - Criteria
    ' Column Q - Ticker
    ' Column R - Value
' ==================================================================================================

    'Loop through all worksheets
    For Each ws In ActiveWorkbook.Worksheets
        
            Range("A1").Select
            Last_Row = 2
            Last_Row_Summary_Table = 2
            Summary_Table_Row = 2

            ' Add Column Headings for Summary Table
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Open Price"
            ws.Range("K1") = "Close Price"
            ws.Range("L1") = "Yearly Change"
            ws.Range("M1") = "Percent Change"
            ws.Range("N1") = "Total Stock Volumn"

            ' Format number/percent to display data
            ws.Range("J:J").NumberFormat = "#,##0.00"
            ws.Range("K:K").NumberFormat = "#,##0.00"
            ws.Range("L:L").NumberFormat = "#,##0.00"
            ws.Range("M:M").NumberFormat = "0.00%"
            ws.Range("N:N").NumberFormat = "#,##0"
            ws.Range("R2").NumberFormat = "0.00%"
            ws.Range("R3").NumberFormat = "0.00%"
            ws.Range("R4").NumberFormat = "#,##0"

            ' Check to see if the next row contains the same stock ticker symbol as the current row.
            ' If the same stock ticker symbol, run the calculations
            ' If different stock ticker symbol, advance to the next row in the Summary Table for next stock ticker....
                
            ' Find Last Row of Ticker data
            Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Loop through all stock trades
            For i = 2 To Last_Row

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    ' Print the Ticker Symbol to the Summary Table
                    Ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & Summary_Table_Row).Value = Ticker

                    ' Print the Close Price of the Stock to the Summary Table
                    Close_Price = ws.Range("F" & i).Value

                    ' Print the Yearly Change to the Summary Table
                    ws.Range("L" & Summary_Table_Row) = Close_Price - Open_Price
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    Yearly_Change = Close_Price - Open_Price
                    If Open_Price = 0 Then
                        Percent_Change = 0
                    Else
                        ws.Range("M" & Summary_Table_Row) = (Close_Price - Open_Price) / Open_Price
                    End If
                    
                    ' Print the Percent Change to the Summary Table                    
                    ws.Range("L" & Summary_Table_Row).Value = Yearly_Change

                    ' Add Column Heading "Total Stock Volumn" and print the Stock Volumn traded to the Summary Table
                    Volumn = Volumn + ws.Cells(i, 7).Value
                    ws.Range("N" & Summary_Table_Row).Value = Volumn

                    ' Conditional formatting for Yearly Change Column (L) - green for positive, red for negative
                    Last_Row_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
                        
                    If ws.Range("L" & Last_Row_Summary_Table).Value >= 0 Then
                        ws.Range("L" & Last_Row_Summary_Table).Interior.ColorIndex = 4
                    Else: ws.Range("L" & Last_Row_Summary_Table).Interior.ColorIndex = 3                    
                    End If

                    ' Add one to the Summary Table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    Previous_Amount = i + 1
                    Volumn = 0
                    Flag = False

                ' If the cell immediatly following a row is the same Stock Ticker Symbol....
                Else
                    If Flag = False Then
                        Open_Price = ws.Cells(i, 3).Value
                        'Cells(i, 10) = Open_Price
                        Flag = True
                    End If
                    
                    Volumn = Volumn + ws.Cells(i, 7).Value

                End If

            Next i

            ' Add Column Headings for Greatest Total Table
            ws.Range("P1") = "Criteria"
            ws.Range("P2") = "Greatest % Increase"
            ws.Range("P3") = "Greatest % Decrease"
            ws.Range("P4") = "Greatest Total Volumn"
            ws.Range("Q1") = "Ticker"
            ws.Range("R1") = "Value"

            ' Find Ticker and Value with Greatest % Increase
            Greatest_Percent_Increase_Value = WorksheetFunction.Max(ws.Range("M2:M" & Last_Row_Summary_Table).Value)
            ws.Range("R2") = Greatest_Percent_Increase_Value
            Greatest_Percent_Increase_Ticker = WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("M1:M" & Last_Row_Summary_Table), 0)
            Greatest_Percent_Increase_Ticker = ws.Range("I" & Greatest_Percent_Increase_Ticker)
            ws.Range("Q2") = Greatest_Percent_Increase_Ticker

            ' Find Ticker and Value with Greatest % Decrease
            Greatest_Percent_Decrease_Value = WorksheetFunction.Min(ws.Range("M2:M" & Last_Row_Summary_Table).Value)
            ws.Range("R3") = Greatest_Percent_Decrease_Value
            Greatest_Percent_Decrease_Ticker = WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("M1:M" & Last_Row_Summary_Table), 0)
            Greatest_Percent_Decrease_Ticker = ws.Range("I" & Greatest_Percent_Decrease_Ticker)
            ws.Range("Q3") = Greatest_Percent_Decrease_Ticker

            ' Find Ticker and Value with Greatest Total Volumn
            Greatest_Total_Volumn_Value = WorksheetFunction.Max(ws.Range("N2:N" & Last_Row_Summary_Table).Value)
            ws.Range("R4") = Greatest_Total_Volumn_Value
            Greatest_Total_Volumn_Ticker = WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("N1:N" & Last_Row_Summary_Table), 0)
            Greatest_Total_Volumn_Ticker = ws.Range("I" & Greatest_Total_Volumn_Ticker)
            ws.Range("Q4") = Greatest_Total_Volumn_Ticker

            ' Autofit to display data and format Columns
            ws.Columns("I:N").AutoFit
            ws.Columns("P:R").AutoFit

            ' Reset the Total Stock Volumn Total, Open Price, Close Price
            Volumn = 0
            Open_Price = 0
            Close_Price = 0

    Next ws

End Sub

