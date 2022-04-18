Sub stockSMU()

Dim ws As Worksheet
For Each ws In Worksheets

        Dim i As Long
        Dim j As Integer
        Dim summaryTable As Integer
        Dim finalRow As Long
        Dim cumVolume As Double
        Dim percYearDelta As Double
        Dim initialOpen As Double
        Dim percDelta As Double
        Dim greatestDeltaIncr As Double
        Dim greatestDeltaIncrDec As Double
        Dim MaxVolume As Double
        
    summaryTable = 2
    cumVolume = 0
    initialOpen = Cells(2, 3).Value
    ws.Cells(1, 9) = "Stock Ticker"
    ws.Cells(1, 10) = "Percentage Year Delta"
    ws.Cells(1, 11) = "Percentage Delta"
    ws.Cells(1, 12) = "Cumulative Volume"
    ws.Cells(1, 15) = "Final Value"

    finalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To finalRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(summaryTable, 9).Value = ws.Cells(i, 1).Value
            cumVolume = cumVolume + ws.Cells(i, 7).Value
            ws.Cells(summaryTable, 12).Value = cumVolume
            
            percYearDelta = ws.Cells(i, 6).Value - initialOpen
            
            If initialOpen > 0 Then
                percDelta = percYearDelta / initialOpen
            Else
                percDelta = -1
            End If
            ws.Cells(summaryTable, 10).Value = percYearDelta
            ws.Cells(summaryTable, 11).Value = percDelta
            
            If percYearDelta < 0 Then
                ws.Cells(summaryTable, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryTable, 10).Interior.ColorIndex = 4
            End If
            
            summaryTable = summaryTable + 1
            cumVolume = 0
            percDelta = 0
            initialOpen = ws.Cells(i + 1, 3).Value
      
        Else
           cumVolume = cumVolume + ws.Cells(i, 7).Value
           
             End If
            
    Next i
    
        ws.Range("K:K").NumberFormat = "0.00%"

        ws.Cells(1, 16) = "Final Value"
        ws.Cells(2, 14) = "Largest Percentage Increase"
        ws.Cells(3, 14) = "Largest Percentage Decrease"
        ws.Cells(4, 14) = "Largest Total Volume"
        
        greatestDeltaIncr = Application.WorksheetFunction.Max(ws.Columns("K"))
        greatestDeltaIncrDec = Application.WorksheetFunction.Min(ws.Columns("K"))
        MaxVolume = Application.WorksheetFunction.Max(ws.Columns("L"))
        
        ws.Cells(2, 16).Value = greatestDeltaIncr
        ws.Cells(3, 16).Value = greatestDeltaIncrDec
        ws.Cells(4, 16).Value = MaxVolume
       
        
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        For k = 2 To finalRow
            If ws.Cells(k, 11).Value = greatestDeltaIncr Then
            ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
            End If
            If ws.Cells(k, 11).Value = greatestDeltaIncrDec Then
            ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
            End If
            If ws.Cells(k, 12).Value = MaxVolume Then
            ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
            End If
        Next k
        

Next

End Sub
