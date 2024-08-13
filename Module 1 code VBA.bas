Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim lastOutput As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim j As Long
    Dim i As Long

    For Each ws In Worksheets
        ' Find the last row of data
        rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        j = 2 ' Initialize the output row counter

        ' Loop through each row of data
        For i = 2 To rowCount
            ' Check if the next row has a different ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                endRow = i

                ' Find the start row for the current ticker
                For startRow = endRow To 2 Step -1
                    If ws.Cells(startRow - 1, 1).Value <> ticker Then
                        Exit For
                    End If
                Next startRow

                ' Calculate the opening price
                openingPrice = ws.Cells(startRow, 3).Value
                ' Calculate the closing price
                closingPrice = ws.Cells(endRow, 6).Value

                ' Calculate the quarterly change
                quarterlyChange = closingPrice - openingPrice
                ' Calculate the percent change
                percentChange = (quarterlyChange / openingPrice) * 100

                ' Calculate the total volume
                totalVolume = 0
                For startRow = startRow To endRow
                    totalVolume = totalVolume + ws.Cells(startRow, 7).Value
                Next startRow

                ' Output the results
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = quarterlyChange
                ws.Cells(j, 11).Value = percentChange
                ws.Cells(j, 12).Value = totalVolume

                j = j + 1 ' Move to the next output row
            End If
        Next i
        
        ' Apply conditional formatting to the Percent Change column
        lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        With ws.Range("K2:K" & lastRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(144, 238, 144) '
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(255, 99, 71) '
            End With
    Next ws
    ' Call the new sub to findgreatest changes
    Call FindGreatestChanges
End Sub



