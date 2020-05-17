Sub StockData()
Dim Ws_Count As Integer
Dim S As Integer

Ws_Count = ActiveWorkbook.Worksheets.Count

For S = 1 To Ws_Count
    
    Dim LastColumn As Integer
    'find last column with data
    LastColumn = Cells(1.1).End(xlToRight).Column
    
    'Create the headers for the summary table
    Cells(1, (LastColumn + 2)).Value = "Ticker"
    Cells(1, (LastColumn + 3)).Value = "Year Change"
    Cells(1, (LastColumn + 4)).Value = "Percent Change"
    Cells(1, (LastColumn + 5)).Value = "Total Stock Volume"
      
    'set the variables for creating the summary table
    Dim Ticker_name As String
    Dim Last_Row_Column_A As Long
    
    'find the last row with data in Column A (Ticker)
    Last_Row_Column_A = Cells(Rows.Count, 1).End(xlDown).Row
           
    'set the variable for the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    Dim Total_Volume As LongLong
    Total_Volume = 0
    
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Price_Change As Double
    Dim Percent_Change As Double
    
         
    'loop for each entry in the Column A (Ticker)
    For i = 2 To Last_Row_Column_A
       'check if we are still within the same ticker in the column A, and if not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'set the ticker name
            Ticker_name = Cells(i, 1).Value
            
            'Add to the total volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            'figure and add the Price_Change (i am sure that the Open Price variable doesn't pick up the right value,
            'but haven't figured how to fix it. It picks up the open price for the last day of the year...
                            
            Close_Price = Cells(i, 6).Value
            
            Open_Price = Cells(i, 3).Value
                         
            Price_Change = Close_Price - Open_Price
                If Open_Price <> 0 Then
                
                    Percent_Change = Price_Change / Open_Price
                Else
                    Percent_Change = 1
                End If
                
            'print the Tickers to the summary table
            Range("I" & Summary_Table_Row).Value = Ticker_name
            
            'print Yearly Change to the Summary Row
            Range("J" & Summary_Table_Row).Value = Price_Change
            
            'print percentage change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            If Range("K" & Summary_Table_Row).Value > 0 Then
                Range("K" & Summary_Table_Row).Interior.Color = vbGreen
            Else
                Range("K" & Summary_Table_Row).Interior.Color = vbRed
            
            End If
                                 
            'print total volume to the summary table
            Range("L" & Summary_Table_Row).Value = Total_Volume
            'add one to the summary row
            Summary_Table_Row = Summary_Table_Row + 1
            'reset Total Volume variable
            Total_Volume = 0
            
        'if the cell immediately following a row within the same ticker
        Else
            'add to the Total_Volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
       
       End If
            
    Next i
Next S

End Sub