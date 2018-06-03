Sub TickerSummary()
    
    'DECLARE VARIABLES
    Dim ws As Worksheet                 'worksheet object
    Dim Worksheet_Name As String        'name of the active worksheet
    Dim Last_Row As Long                'last row of the active worksheet
    Dim Summary_Table_Row As Long       'row number of the summary table
    Dim Ticker_Name As String           'ticker name
    Dim Total_Stock_Volume As Double    'cumulative stock volume for each ticker
    Dim Year_Open As Double             'ticker's year opening price
    Dim Yearly_Change As Double         'annual price change
    Dim Percent_Change As Variant       'annual price change %
    Dim GPI_Ticker As String            'holds ticker name with greatest % increase
    Dim GPI_Value As Variant            'holds value of greatest % increase
    Dim GPD_Ticker As String            'holds ticker name with greatest % decrease
    Dim GPD_Value As Variant            'holds value of greatest % decrease
    Dim GTV_Ticker As String            'holds ticker name with greatest total volume
    Dim GTV_Value As Double             'holds value of greatest total volume
        
    
    'LOOP THROUGH ALL THE WORKSHEETS
    
    For Each ws In Worksheets
    
        'Determine Last Row of Active Worksheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sort Rows in Ascending Order by Ticker Value and Date
        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(Last_Row, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ws.Sort.SortFields.Add Key:=Range(Cells(2, 2), Cells(Last_Row, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ws.Sort
            .SetRange Range(Cells(2, 1), Cells(Last_Row, 7))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
        'Initialize the Variables for Active Worksheet
        Worksheet_Name = ws.Name
        Summary_Table_Row = 2
        Ticker_Name = ws.Cells(2, 1).Value
        Total_Stock_Volume = 0
        Year_Open = ws.Cells(2, 3).Value
        GPI_Ticker = ws.Cells(2, 1).Value
        GPI_Value = 0
        GPD_Ticker = ws.Cells(2, 1).Value
        GPD_Value = 0
        GTV_Ticker = ws.Cells(2, 1).Value
        GTV_Value = 0
    
        'Format Summary table for Active Worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Format percentage cells and columns
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
        For i = 2 To Last_Row
        
            'If ticker changes in the next row, add last volume entry to total and record results in summary table
            If ws.Cells(i + 1, 1).Value <> Ticker_Name Then
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                Yearly_Change = ws.Cells(i, 6).Value - Year_Open
                
                'Account for annual open price of zero (prevents divided by zero error)
                If Year_Open <> 0 Then
                    Percent_Change = Yearly_Change / Year_Open
                
                Else
                    Percent_Change = "N/A"
                
                End If
                
                'Record values to summary table
                ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                
                'Color format yearly change value
                If (ws.Cells(i, 6).Value - Year_Open) > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 10
                    
                Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Update greatest % increase ticker and value
                If Percent_Change > GPI_Value And Percent_Change <> "N/A" Then
                    GPI_Ticker = Ticker_Name
                    GPI_Value = Percent_Change
                End If
                
                'Update greatest % decrease ticker and value
                If Percent_Change < GPD_Value And Percent_Change <> "N/A" Then
                    GPD_Ticker = Ticker_Name
                    GPD_Value = Percent_Change
                End If
                
                'Update greatest total volume ticker and value
                If Total_Stock_Volume > GTV_Value Then
                    GTV_Ticker = Ticker_Name
                    GTV_Value = Total_Stock_Volume
                End If
                
                'Increment summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset ticker name and ticker summary values
                Ticker_Name = ws.Cells(i + 1, 1).Value
                Total_Stock_Volume = 0
                Yearly_Change = 0
                Percent_Change = 0
                Year_Open = ws.Cells(i + 1, 3).Value
                
                
            'Otherwise add stock volume entry to current total
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'Record greatest values and tickers to summary table
        ws.Cells(2, 16).Value = GPI_Ticker
        ws.Cells(2, 17).Value = GPI_Value
        ws.Cells(3, 16).Value = GPD_Ticker
        ws.Cells(3, 17).Value = GPD_Value
        ws.Cells(4, 16).Value = GTV_Ticker
        ws.Cells(4, 17).Value = GTV_Value
        
        'Final formatting of summary table
        ws.Columns.AutoFit
    
    Next ws

End Sub

