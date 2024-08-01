Attribute VB_Name = "Module2"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim tickerIncrease As String, tickerDecrease As String, tickerVolume As String
    Dim rng As Range

    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet has data in it
        If WorksheetFunction.CountA(ws.Cells) > 0 Then
            ' Find the last row with data in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Insert new columns
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % increase"
            ws.Cells(3, 15).Value = "Greatest % decrease"
            ws.Cells(4, 15).Value = "Greatest total volume"

            ' Initialize variables for tracking maximum values
            maxIncrease = -1E+30 ' Very small number
            maxDecrease = 1E+30 ' Very large number
            maxVolume = -1E+30 ' Very small number

            ' Loop through each row of data to find greatest values
            For i = 2 To lastRow
                ' Check for greatest % increase
                If ws.Cells(i, 11).Value > maxIncrease Then
                    maxIncrease = ws.Cells(i, 11).Value
                    tickerIncrease = ws.Cells(i, 9).Value
                End If
                
                ' Check for greatest % decrease
                If ws.Cells(i, 11).Value < maxDecrease Then
                    maxDecrease = ws.Cells(i, 11).Value
                    tickerDecrease = ws.Cells(i, 9).Value
                End If
                
                ' Check for greatest total volume
                If ws.Cells(i, 12).Value > maxVolume Then
                    maxVolume = ws.Cells(i, 12).Value
                    tickerVolume = ws.Cells(i, 9).Value
                End If
            Next i
            
            ' Output results
            ws.Cells(2, 16).Value = tickerIncrease
            ws.Cells(2, 17).Value = maxIncrease
            
            ws.Cells(3, 16).Value = tickerDecrease
            ws.Cells(3, 17).Value = maxDecrease
            
            ws.Cells(4, 16).Value = tickerVolume
            ws.Cells(4, 17).Value = maxVolume
            
        End If
    Next ws
    
    MsgBox "Stock analysis complete!", vbInformation
End Sub



