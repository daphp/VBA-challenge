Attribute VB_Name = "Module1"
Sub ProcessStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range
    Dim cell As Range
    Dim dateValue As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String

    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet has data in it
        If WorksheetFunction.CountA(ws.Cells) > 0 Then
            ' Find the last row with data in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Convert dates in column B
            Set rng = ws.Range("B1:B" & lastRow)
            For Each cell In rng
                If IsNumeric(cell.Value) And Len(cell.Value) = 8 Then
                    dateValue = cell.Value
                    yearPart = Left(dateValue, 4)
                    monthPart = Mid(dateValue, 5, 2)
                    dayPart = Right(dateValue, 2)
                    cell.Value = monthPart & "/" & dayPart & "/" & yearPart
                    cell.NumberFormat = "mm/dd/yyyy"
                End If
            Next cell

            ' Insert new columns
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            ' Loop through each row of data
            For i = 2 To lastRow
                ' Copy ticker value
                ws.Cells(i, 9).Value = ws.Cells(i, 1).Value

                ' Calculate quarterly change
                ws.Cells(i, 10).Value = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value

                ' Calculate percent change
                If ws.Cells(i, 3).Value <> 0 Then
                    ws.Cells(i, 11).Value = (ws.Cells(i, 10).Value / ws.Cells(i, 3).Value) * 100
                Else
                    ws.Cells(i, 11).Value = 0
                End If

                ' Copy total stock volume
                ws.Cells(i, 12).Value = ws.Cells(i, 7).Value
            Next i

            ' Apply conditional formatting to the "Quarterly Change" column (column 10)
            Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10))

            ' Clear previous formatting
            rng.FormatConditions.Delete

            ' Format negative values in red
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                .Interior.Color = RGB(255, 0, 0) ' Red
            End With

            ' Format positive values in green
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
                .Interior.Color = RGB(0, 255, 0) ' Green
            End With
        End If
    Next ws

    MsgBox "Data processing complete!", vbInformation
End Sub

