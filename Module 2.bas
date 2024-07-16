Attribute VB_Name = "Module2"
Sub FindGreatestChanges()
    Dim ws As Worksheet
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim lastRow As Long
    Dim i As Long
    
    
    maxIncrease = -99999
    maxDecrease = 99999
    maxVolume = 0
    
    ' Asumming output is on first worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
   ' Find the last row of the output data
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    ' Loop through the output data to find the greatest changes
    For i = 2 To lastRow
        If ws.Cells(i, 11).Value > maxIncrease Then
            maxIncrease = ws.Cells(i, 11).Value
            maxIncreaseTicker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < maxDecrease Then
            maxDecrease = ws.Cells(i, 11).Value
            maxDecreaseTicker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > maxVolume Then
            maxVolume = ws.Cells(i, 12).Value
            maxVolumeTicker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ' Output the greatest changes
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = maxIncreaseTicker
    ws.Cells(2, 16).Value = maxIncrease / 100
    ws.Cells(2, 16).NumberFormat = "0.00%"
    
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = maxDecreaseTicker
    ws.Cells(3, 16).Value = maxDecrease / 100
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
    
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = maxVolumeTicker
    ws.Cells(4, 16).Value = maxVolume
End Sub

