Attribute VB_Name = "Module1"
Sub CalculateStockMetricsAcrossSheets()
    ' Declare variables for holding the maximums and minimums
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim startRow As Integer
    Dim ticker As String
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim i As Integer
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize the maximum and minimum values for each worksheet
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        startRow = 2 ' Assume data starts at row 2
        
        ' Loop through each row of data
        For i = startRow To lastRow
            ' Check if this is the first row of the current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                yearOpen = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Add the volume to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if this is the last row of the current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                yearClose = ws.Cells(i, 6).Value
                yearlyChange = yearClose - yearOpen
                If yearOpen <> 0 Then ' Prevent division by zero
                    percentChange = (yearlyChange / yearOpen) * 100
                Else
                    percentChange = 0
                End If
                
                ' Check for max increase, max decrease, and max volume
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                ElseIf percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
            End If
        Next i
        
        ' Place the summary results at the top of the worksheet
        With ws
            .Cells(1, 9).Value = "Greatest % Increase"
            .Cells(1, 10).Value = maxIncreaseTicker
            .Cells(1, 11).Value = maxIncrease
            
            .Cells(2, 9).Value = "Greatest % Decrease"
            .Cells(2, 10).Value = maxDecreaseTicker
            .Cells(2, 11).Value = maxDecrease
            
            .Cells(3, 9).Value = "Greatest Total Volume"
            .Cells(3, 10).Value = maxVolumeTicker
            .Cells(3, 11).Value = maxVolume
        End With
        
        ' Apply conditional formatting to the percentage change column
        Dim rng As Range
        Set rng = ws.Range("G" & startRow & ":G" & lastRow) ' Assuming percentage change is in column G
        
        ' Clear any previous conditional formatting
        rng.FormatConditions.Delete
        
        ' Add conditional formatting for positive change
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0) ' Green
        End With
        
        ' Add conditional formatting for negative change
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Red
        End With
    Next ws
End Sub

