' VBA Homework - The VBA of Wall Street
' ----------------------------------------------------------------------------
'Instructions:
'Create a script that will loop through all the stocks for one year and output the following information:
'   - The ticker symbol.
'   - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   - The total stock volume of the stock.
'   - You should also have conditional formatting that will highlight positive change in green and negative change in red.
'------------------------------------------------------------------------------
'BONUS:
' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
' Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
'------------------------------------------------------------------------------


Sub StockAnalysis()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

    'Inserting Summary Titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Open_Price"
    ws.Range("K1").Value = "Close_Price"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percentage Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("Q2").Value = "Greatest % Increase"
    ws.Range("Q3").Value = "Greatest % Decrease"
    ws.Range("Q4").Value = "Greatest Total Volume"
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"
    
    ' Flag
    Dim Flag As Integer
    
    ' Set an initial variable for holding the ticker symbol
    Dim Ticker As String
    
    ' Set variables to calculate Yearly Change per Ticker
    Dim Initial_Value As Double
    Dim Final_Value As Double
    
    'Set variables for summary table
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    'Set variables for bonus
    Dim Max_Change As Double
    Dim Min_Change As Double
    Dim Max_Vol As Double
    
    Dim Max_Ticker As String
    Dim Min_Ticker As String
    Dim MaxV_Ticker As String
    
    ' Keep track of the location for ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Count the last and first row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Flag = 0

    
        ' Loop through each row
        For i = 2 To lastRow
            
            If Flag = 0 Then
                
                'Find open price per ticker
                Initial_Value = ws.Cells(i, 3).Value
                Flag = 1
            End If
                 
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                
                   
            'Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
           
                
            End If

            
            'Check if we are still within the same ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            Flag = 0
        
            'Set the ticker symbol
            Ticker = ws.Cells(i, 1).Value

                
            'Find close price per ticker
            Final_Value = ws.Cells(i, 6).Value
                
            'Calculate yearly change per ticker
            Yearly_Change = Final_Value - Initial_Value
                
                'Calculate percent change per ticker and
                If Initial_Value = 0 Then
                        
                    Percent_Change = 0
                    
                Else
                        
                    Percent_Change = Yearly_Change / Initial_Value
                    
                End If
                    
    
            'Add to the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            'Print Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Initial_Value
            ws.Range("K" & Summary_Table_Row).Value = Final_Value
            ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("M" & Summary_Table_Row).Value = Percent_Change
            ws.Range("N" & Summary_Table_Row).Value = Total_Stock_Volume
        
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Reset the Total Stock Volume
            Total_Stock_Volume = 0
            
            End If
               
        Next i
        
    'Loop to format positive change in green and negative in red
    For j = 2 To Summary_Table_Row
    
        If ws.Cells(j, 12).Value < 0 Then
            ws.Cells(j, 12).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 12).Interior.ColorIndex = 4
        End If

    Next j
    
    
    'Bonus: retrieve greatest % increase, % decrease and total volume
    lastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
    Max_Change = ws.Cells(2, 13).Value
    Min_Change = ws.Cells(2, 13).Value
    Max_Vol = ws.Cells(2, 14).Value
    Max_Ticker = ws.Cells(2, 9).Value
    Min_Ticker = ws.Cells(2, 9).Value
    MaxV_Ticker = ws.Cells(2, 9).Value
    
    For x = 2 To lastRowSummary
        
        If ws.Cells(x, 13).Value > Max_Change Then
            Max_Change = ws.Cells(x, 13).Value
            Max_Ticker = ws.Cells(x, 9).Value
        End If
        
        If ws.Cells(x, 13).Value < Min_Change Then
            Min_Change = ws.Cells(x, 13).Value
            Min_Ticker = ws.Cells(x, 9).Value
        End If
        
        If ws.Cells(x, 14).Value > Max_Vol Then
            Max_Vol = ws.Cells(x, 14).Value
            MaxV_Ticker = ws.Cells(x, 9).Value
        End If
    
    'Print Summary Table
    Next x
    
    ws.Cells(2, 18) = Max_Ticker
    ws.Cells(2, 19) = Max_Change
    
    ws.Cells(3, 18) = Min_Ticker
    ws.Cells(3, 19) = Min_Change
    
    ws.Cells(4, 18) = MaxV_Ticker
    ws.Cells(4, 19) = Max_Vol
    
    'Add format to percentages and columns width
    ws.Columns("I:N").AutoFit
    ws.Columns("Q:S").AutoFit
    ws.Range("M:M").NumberFormat = "0.00%"
    ws.Range("S2:S3").NumberFormat = "0.00%"

    Next ws
       
End Sub
