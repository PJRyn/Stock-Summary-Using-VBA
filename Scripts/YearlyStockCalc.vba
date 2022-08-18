Attribute VB_Name = "YearlyStockCalc"
Sub YearlyStockCalc():
    'Code to go through each worksheet within a file
    For Each ws In Worksheets
        Dim WorksheetName As String
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        WorksheetName = ws.Name
        'Tells the user which worksheet is being loaded
        MsgBox (WorksheetName)
        
        'initialise variables used:
        Dim TickName As String
        'Counter for creating the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        'Open and closing values and calculations for them
        Dim OpenVal As Double
        Dim CloseVal As Double
        Dim YearSum As Double
        Dim YearPrcnt As Double
        Dim VolCount As Double
        
        'Collect the first entries Year opening value
        OpenVal = ws.Range("C2").Value
        
        'Create Summary Table
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        'count through the list of rows checking for new Tickers
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the Ticker name
                TickName = ws.Cells(i, 1).Value
                ' Print the Credit Card Brand in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = TickName
                
                'Collect Close Value and Calculate a Yearly difference between yearly opening and closing value
                CloseVal = ws.Cells(i, 6).Value
                YearSum = CloseVal - OpenVal
                
                'Format Yearly Summary
                ws.Range("K" & Summary_Table_Row).Value = YearSum
                ws.Range("K" & Summary_Table_Row).NumberFormat = "$#,##0.00"
                If YearSum > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf YearSum < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                ElseIf YearSum = 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 6
                End If
                
                'Calculate and print percentage change
                YearPrcnt = (CloseVal - OpenVal) / CloseVal
                ws.Range("L" & Summary_Table_Row).Value = YearPrcnt
                ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ws.Range("M" & Summary_Table_Row).Value = VolCount
                VolCount = 0
                
                'Collect new opening value
                OpenVal = ws.Cells(i + 1, 3).Value
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            Else
        
              ' Add to the Brand Total
              VolCount = VolCount + ws.Cells(i, 7).Value
            
            End If
        Next i
    Next ws
    
    MsgBox ("Completed!")
    
End Sub
        
