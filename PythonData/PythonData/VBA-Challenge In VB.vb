Sub VBA_challenge()

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

' Declare variables

' Set variable for the ticker symbol
Dim Ticker As String

' Set variable for the opening stock price
Dim Open_Price As Double

' Set variable for the closing stock price
Dim Close_Price As Double

' Set variable for the stock trading volumn and initial number traded for counter
Dim Volume As Double
Volume = 0

' Keep track of the location for each Ticker Symbol in the Summary Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Set variable to find last row of Ticker data
Dim Last_Row As Long

' Set variable to find last row of Summary Table data
Dim Last_Row_Summary_Table As Long

' Set variable for the Yearly Change number in Summary table
Dim Yearly_Change As Double

' Set variable for Worksheet
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

' Set variable for Greatest % Increase Value
Dim Greatest_Percent_Increase_Value As Double

' Set variable for Greatest % Decrease Value
Dim Greatest_Percent_Decrease_Value As Double

' Set variable for Greatest Total Volume Ticker
Dim Greatest_Total_Volumn_Ticker As String

' Set variable for Greatest Total Volume Value
Dim Greatest_Total_Volumn_Value As Double

' Set
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
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
            Range("A1").Select
            Last_Row = 2
            Last_Row_Summary_Table = 2
            Summary_Table_Row = 2
            
            ' Loop through all stock trades
            ' Find Last Row of Ticker data
            
            ' LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To Last_Row
            
                ' Check to see if the next row contains the same stock ticker symbol as the current row.
                ' If not, advance to the next row in the Summary Table....
                
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    ' Set the new stock ticker symbol name
                    Ticker = Cells(i, 1).Value
                    
                    ' Add the volumn traded of the stock to volumn
                    Volumn = Volumn + Cells(i, 7).Value
                    
                    ' Add Column Heading "Ticker" and print the Ticker Symbol to the Summary Table
                    Range("I1").Value = "Ticker"
                    Range("I" & Summary_Table_Row).Value = Ticker
                                
                    ' Add Column Heading "Open Price" and print the Open Price to the Summary Table
                    Range("J1") = "Open Price"
                    ' Range("J" & Summary_Table_Row).Value = Open_Price
                    ' Open_Price = Open_Price + Cells(i, 3).Value
                    
                    
                    ' Add Column Heading "Close Price" and print the Open Price to the Summary Table
                    Range("K1") = "Close Price"
            
                                   
                    ' Add Column Heading "Yearly Change" and print the Open Price to the Summary Table
                    Range("L1") = "Yearly Change"
                    ' Calculate Yearly Change  ===Morgan
                    Close_Price = ws.cells(i, 6).value - Open_Price
                                  
                    ' Add Column Heading "Percent Change" and print the Open Price to the Summary Table
                    Range("M1") = "Percent Change"
                                    
                    ' Add Column Heading "Total Stock Volumn" and print the Stock Volumn traded to the Summary Table
                    Range("N1") = "Total Stock Volumn"
                    Range("N" & Summary_Table_Row).Value = Volumn
                                    
                    ' Add one to the Summary Table row
                    Summary_Table_Row = Summary_Table_Row + 1
                                
                    '  RESET HERE ===========================================================================
                    
                ' If the cell immediatly following a row is the same Stock Ticker Symbol....
                Else
                
                    ' Add to the Total Stock Volumn Total
                    Volumn = Volumn + Cells(i, 7).Value
            

            
                ' Find last row of Summary Table
                Last_Row_Summary_Table = Cells(Rows.Count, 9).End(xlUp).Row
                
                ' Calculate the Yearly Change
                Range("L2:L" & Last_Row_Summary_Table).Formula = "=K2-J2"
                Range("L2:L" & Last_Row_Summary_Table).Value = Range("L2:L" & Last_Row_Summary_Table).Value
                
                ' Calculate the Percent Change
                Range("M2:M" & Last_Row_Summary_Table).Formula = "=K2/(K2-J2)"
                Range("M2:M" & Last_Row_Summary_Table).Value = Range("M2:M" & Last_Row_Summary_Table).Value
                
                'Conditional formatting for Yearly Change Column (L) - green for positive, red for negative
                For Each Cell In Range("L2:L" & Last_Row_Summary_Table)
                    If Cell.Value > 0 Then
                    Cell.Offset(0, 0).Interior.ColorIndex = 4
                    ElseIf Cell.Value < 0 Then
                    Cell.Offset(0, 0).Interior.ColorIndex = 3
                    End If
                Next Cell

            
                    
        ' Add Column Heading "Criteria" to the Greatest Stocks Table
        Range("P1") = "Criteria"
        
        ' Add Column Heading "Ticker" to the Greatest Stocks Table
        Range("Q1") = "Ticker"
        
        ' Add Column Heading "Value" to the Greatest Stocks Table
        Range("R1") = "Value"
            
        ' Add Row Heading "Greatest % Increase" to the Greatest Stocks Table
        Range("P2") = "Greatest % Increase"
        
        ' Find Ticker and Value with Greatest % Increase
        Greatest_Percent_Increase_Value = WorksheetFunction.Max(Range("I2:I" & Last_Row_Summary_Table).Value)
        Range("R2") = Greatest_Percent_Increase_Value
        ' Greatest_Percent_Increase_Ticker = WorksheetFunction.Match(Range("R2").Value, Range("M1:M" & Last_Row_Summary_Table), 0)
        ' Greatest_Percent_Increase_Ticker = Range("I" & Greatest_Percent_Increase_Ticker)
        
        ' Add Row Heading "Greatest % Decrease" to the Greatest Stocks Table
        Range("P3") = "Greatest % Decrease"
        
        ' Find Ticker and Value with Greatest % Decrease
        Greatest_Percent_Decrease_Value = WorksheetFunction.Min(Range("I2:I" & Last_Row_Summary_Table).Value)
        Range("R3") = Greatest_Percent_Decrease_Value
        Range("Q3") = Greatest_Percent_Decrease_Ticker
        ' Greatest_Percent_Decrease_Ticker = WorksheetFunction.Match(Range("R3").Value, Range("M1:M" & Last_Row_Summary_Table), 0)
        ' Greatest_Percent_Decrease_Ticker = Range("I" & Greatest_Percent_Decrease_Ticker)
        
        ' Add Row Heading "Greatest Total Volumn" to the Greatest Stocks Table
        Range("P4") = "Greatest Total Volumn"
                    
        ' Find Ticker and Value with Greatest Total Volumn
        Greatest_Total_Volumn_Value = WorksheetFunction.Max(Range("N2:N" & Last_Row_Summary_Table).Value)
        Range("R4") = Greatest_Total_Volumn_Value
        Range("Q4") = Greatest_Total_Volumn_Ticker
        Greatest_Total_Volumn_Ticker = WorksheetFunction.Match(Range("R4").Value, Range("N1:N" & Last_Row_Summary_Table), 0)
        Greatest_Total_Volumn_Ticker = Range("I" & Greatest_Total_Volumn_Ticker)
        
        
        ' Autofit to display data and format Columns
        Range("J:J").NumberFormat = "#,##0.00"
        Range("K:K").NumberFormat = "#,##0.00"
        Range("L:L").NumberFormat = "#,##0"
        Range("M:M").NumberFormat = "0.00%"
        Range("N:N").NumberFormat = "#,##0"
        Range("R2").NumberFormat = "0.00%"
        Range("R3").NumberFormat = "0.00%"
        Range("R4").NumberFormat = "#,##0"
        Columns("I:N").AutoFit
        Columns("P:R").AutoFit
        
        End If
            
            ' reset everything
            ' Reset the Total Stock Volumn Total
            Volumn = 0
            Open_Price = 0
            Close_Price = 0

             
            Flag = False
            
        If Flag = False Then
        Open_Price = ws.Cells(i, 3).Value
        Flag = True
        End If
                     
        Next i
        
    Next ws
    
starting_ws.Activate
        
End Sub

