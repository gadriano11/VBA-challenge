Attribute VB_Name = "Module1"
Sub Stock_Checker():

    For Each ws In Worksheets
        'Declare variables
        Dim ws_Name As String
        Dim Last_Row_A As Long
        Dim Last_Row_I As Long
        Dim result As Long
        Dim Percent_Change As Double
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim i As Long
        Dim j As Long

        
        'WorksheetName
        ws_Name = ws.Name
        
        'Column Headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'first row for result and data samples
        result = 2
        j = 2

        'Set last row for ticker
        Last_Row_A = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows for summary results
            For i = 2 To Last_Row_A
                'Checks if still within the same ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Write ticker code in column I
                    ws.Cells(result, 9).Value = ws.Cells(i, 1).Value
                    
                    'Calculate yearly Change in column J
                    ws.Cells(result, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Change cell color if positive_green)/negative(red)
                    If ws.Cells(result, 10).Value < 0 Then
                    'negative value then red
                    ws.Cells(result, 10).Interior.ColorIndex = 3
                    Else
                    'positive value then green
                    ws.Cells(result, 10).Interior.ColorIndex = 4
                    End If
                    
                    'Calculate percent change in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                        Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                        'Percent formating
                        ws.Cells(result, 11).Value = Format(Percent_Change, "Percent")
                    Else
                        ws.Cells(result, 11).Value = Format(0, "Percent")
                    End If
                    
                    'Calculate total volume in column L
                    ws.Cells(result, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                    'Next row for ticker and result column
                    result = result + 1
                    j = i + 1
                
                End If
            
            Next i
           
        'Bonus Round
        
        'Set last row for resulting column
        Last_Row_I = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Assign Variables
        Greatest_Increase = ws.Cells(2, 11).Value
        Greatest_Decrease = ws.Cells(2, 11).Value
        Greatest_Volume = ws.Cells(2, 12).Value
        
            'Loop through the rows for greatest values column
            For i = 2 To Last_Row_I
            
                'Calculate for greatest increase
                If ws.Cells(i, 11).Value > Greatest_Increase Then
                    Greatest_Increase = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                    Greatest_Increase = Greatest_Increase
                End If
            
                'Calculate for greatest decrease
                If ws.Cells(i, 11).Value < Greatest_Decrease Then
                    Greatest_Decrease = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                    Greatest_Decrease = Greatest_Decrease
                End If
            
                'Calculate for greatest total volume
                If ws.Cells(i, 12).Value > Greatest_Volume Then
                    Greatest_Volume = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                    Greatest_Volume = Greatest_Volume
                End If
                       
                'greatest value results with formatting
                ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
                ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
                ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
            
            Next i
            
        'Adjust column width
        Worksheets(ws_Name).Columns("A:Z").AutoFit

    Next ws
        
End Sub
