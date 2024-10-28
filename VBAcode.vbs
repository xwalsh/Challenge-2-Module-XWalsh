Sub StockDataAnalysis()

    Dim sheet As Worksheet
    Dim lastDataRow As Long
    Dim stockTicker As String
    Dim initialPrice As Double
    Dim finalPrice As Double
    Dim changeInQuarter As Double
    Dim changePercentage As Double
    Dim cumulativeVolume As Double
    Dim outputRow As Long
    Dim index As Long, rowIndex As Long, volumeIndex As Long

    ' Loop through all worksheets
    For Each sheet In ActiveWorkbook.Worksheets

        sheet.Activate

        ' Identify the last row with data
        lastDataRow = sheet.Cells(Rows.Count, 1).End(xlUp).row

        ' Initialize header labels
        With sheet
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With

        initialPrice = sheet.Cells(2, 3).Value  ' Column C

        cumulativeVolume = 0
        outputRow = 2

        ' Process each row to gather stock data
        For index = 2 To lastDataRow

            If sheet.Cells(index + 1, 1).Value <> sheet.Cells(index, 1).Value Then
                ' Get ticker symbol
                stockTicker = sheet.Cells(index, 1).Value
                sheet.Cells(outputRow, 9).Value = stockTicker

                ' Closing price and calculations
                finalPrice = sheet.Cells(index, 6).Value  ' Column F
                changeInQuarter = finalPrice - initialPrice
                sheet.Cells(outputRow, 10).Value = changeInQuarter

                ' Calculate percent change
                changePercentage = changeInQuarter / initialPrice
                sheet.Cells(outputRow, 11).Value = changePercentage
                sheet.Cells(outputRow, 11).NumberFormat = "0.00%"

                ' Calculate cumulative volume
                cumulativeVolume = cumulativeVolume + sheet.Cells(index, 7).Value  ' Column G
                sheet.Cells(outputRow, 12).Value = cumulativeVolume

                ' Move to the next output row
                outputRow = outputRow + 1

                ' Update initial price for the next ticker and reset cumulative volume
                initialPrice = sheet.Cells(index + 1, 3).Value
                cumulativeVolume = 0

            Else
                cumulativeVolume = cumulativeVolume + sheet.Cells(index, 7).Value  ' Column G
            End If
        Next index

        ' Apply color coding based on quarterly changes
        Dim lastChangeRow As Long
        lastChangeRow = sheet.Cells(Rows.Count, 9).End(xlUp).row

        For rowIndex = 2 To lastChangeRow
            If sheet.Cells(rowIndex, 10).Value >= 0 Then
                sheet.Cells(rowIndex, 10).Interior.ColorIndex = 10  ' Green for positive changes
            Else
                sheet.Cells(rowIndex, 10).Interior.ColorIndex = 3   ' Red for negative changes
            End If
        Next rowIndex

        ' Set additional headers for analysis
        With sheet
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            .Cells(2, 15).Value = "Highest % Increase"
            .Cells(3, 15).Value = "Highest % Decrease"
            .Cells(4, 15).Value = "Highest Total Volume"
        End With

        ' Find max/min values for each category
        For volumeIndex = 2 To lastChangeRow
            If sheet.Cells(volumeIndex, 11).Value = Application.WorksheetFunction.Max(sheet.Range("K2:K" & lastChangeRow)) Then
                sheet.Cells(2, 16).Value = sheet.Cells(volumeIndex, 9).Value
                sheet.Cells(2, 17).Value = sheet.Cells(volumeIndex, 11).Value
                sheet.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf sheet.Cells(volumeIndex, 11).Value = Application.WorksheetFunction.Min(sheet.Range("K2:K" & lastChangeRow)) Then
                sheet.Cells(3, 16).Value = sheet.Cells(volumeIndex, 9).Value
                sheet.Cells(3, 17).Value = sheet.Cells(volumeIndex, 11).Value
                sheet.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf sheet.Cells(volumeIndex, 12).Value = Application.WorksheetFunction.Max(sheet.Range("L2:L" & lastChangeRow)) Then
                sheet.Cells(4, 16).Value = sheet.Cells(volumeIndex, 9).Value
                sheet.Cells(4, 17).Value = sheet.Cells(volumeIndex, 12).Value
            End If
        Next volumeIndex

        ' Format the output for better visibility
        With sheet.Range("I:Q")
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With

        Worksheets("Q1").Select

    Next sheet

End Sub
