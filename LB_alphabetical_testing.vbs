Attribute VB_Name = "Module1"
Sub WallStreet():
'Create a script that will loop through all the stocks for one year and output the following information:
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    '-------------------------------------------------
    
   'loop all worksheets (taken from Wells Fargo assignment)
    
    Dim ws As Worksheet
        For Each ws In Worksheets
    
    'determine the last row? (taken from Wells Fargo assignment)
    
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        'headers, based on moderate_solution.png in resources (append ws)
            ws.Cells(1, "I").Value = "Ticker"
            ws.Cells(1, "J").Value = "Yearly change"
            ws.Cells(1, "K").Value = "Percent change"
            ws.Range("L1").Value = "Total stock volume"
            
        'defining variables i will need: tickers, open/close dates,yearly change & percent change (from solved image), volume
    
            Dim ticker As String
            Dim opening As Double
            Dim closing As Double
            Dim yearly_change As Double
            Dim percent_change As Double
            Dim volume As LongLong
                volume = 0
            Dim i As LongLong
            Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2

        opening = ws.Cells(2, 3).Value
         
         'loop all tickers
         
        
        For i = 2 To LastRow

        volume = volume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        opening = Cells(Summary_Table_Row, 3)
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'what is the stock?
                    ticker = ws.Cells(i, 1).Value
                    ws.Cells(Summary_Table_Row, 9).Value = ticker
                
                'what is the volume of each stock?
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Cells(Summary_Table_Row, 12).Value = volume
                
                ' resete the volume of each stock.
                    volume = 0
                
                'what is the close cost?
                    closing = ws.Cells(i, 6).Value
                
                'yearly change?
                    opening = ws.Cells(2, 3).Value
                    closing = ws.Cells(i, 6).Value
                    yearly_change = (closing - opening)
                    ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                
                'percent change?
                
                If opening = 0 Then
                    percent_change = 0
                Else
                    percent_change = (yearly_change / opening)
                End If
                    ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                    ws.Cells(Summary_Table_Row, 11).Value = percent_change
                    
                    'Source for number format method: https://www.educba.com/vba-number-format/
                
                ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                
                'reset the opening for each diff stock
                    opening = ws.Cells(i + 1, 3)
        
            'if cells are the same ticker
                Else
                    volume = volume + ws.Cells(i, 7).Value
            End If
            
        Next i

            'From ask_the_class in slack:
            
            LastSummaryRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
             For i = 2 To LastSummaryRow
                 
                 If ws.Cells(i, 10).Value >= 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                 ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                 End If
            Next i
    
    'autofit columns
        ws.Columns("J:L").AutoFit
        

 'remember for;next for all loops, order here matters so next ws is last
    Next ws

    
End Sub
