Sub StockAnalysis()
    ' Declare worksheet object outside the loop
    Dim ws As Worksheet

    ' Variables for greatest values
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolume As Double
    Dim TickerMaxPercentIncrease As String
    Dim TickerMaxPercentDecrease As String
    Dim TickerMaxTotalVolume As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim SummaryRow As Long
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim TickerStartRow As Long
        
        ' Set column headers in the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Find the last row of data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set initial summary table row
        SummaryRow = 2
        ' Initialize TotalVolume here
        TotalVolume = 0
        ' Initialize TickerStartRow
        TickerStartRow = 2

       ' Loop through all rows of data
        For i = 2 To LastRow
    ' Check if the stock volume is numeric before adding it
    If IsNumeric(ws.Cells(i, 7).Value) Then
        ' Add the stock volume to the total volume
        TotalVolume = TotalVolume + CDbl(ws.Cells(i, 7).Value)
    Else
        MsgBox "Non-numeric value encountered in stock volume at row " & i
    End If
            ' Check if the ticker symbol has changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                ' Get the opening price (first occurrence in the group)
                OpeningPrice = ws.Cells(TickerStartRow, 3).Value
                
                ' Get the closing price (last occurrence in the group)
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change
                YearlyChange = ClosingPrice - OpeningPrice
                
                ' Calculate the percent change with correct formatting
                    If OpeningPrice <> 0 Then
                    PercentChange = (ClosingPrice - OpeningPrice) / Abs(OpeningPrice)
                Else
                    PercentChange = 0
                End If

                ' Add the yearly change, percent change, and total volume to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = Format(PercentChange, "0.00%")
                ws.Cells(SummaryRow, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(TickerStartRow, 7), ws.Cells(i, 7)))
                

                ' Update greatest values
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    TickerMaxPercentIncrease = Ticker
                End If
                
                If PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    TickerMaxPercentDecrease = Ticker
                End If
                
                If WorksheetFunction.Sum(ws.Range(ws.Cells(TickerStartRow, 7), ws.Cells(i, 7))) > MaxTotalVolume Then
                    MaxTotalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(TickerStartRow, 7), ws.Cells(i, 7)))
                    TickerMaxTotalVolume = Ticker
                End If

                ' Format the percent change as a percentage
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting for positive and negative yearly changes
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Increment the summary table row
                SummaryRow = SummaryRow + 1
                
                ' Reset the total volume
                TotalVolume = 0
                ' Update TickerStartRow
                TickerStartRow = i + 1
            End If
        Next i
    Next ws

    Dim resultSheet As Worksheet
    On Error Resume Next
    Set resultSheet = Sheets("Greatest Values")
    On Error GoTo 0
    
    If resultSheet Is Nothing Then
        Set resultSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        resultSheet.Name = "Greatest Values"
    End If
    
    resultSheet.Cells.Clear  ' Clear existing content
    
    resultSheet.Cells(1, 1).Value = "Greatest % Increase"
    resultSheet.Cells(2, 1).Value = "Ticker"
    resultSheet.Cells(2, 2).Value = "Value"
    resultSheet.Cells(3, 1).Value = TickerMaxPercentIncrease
    resultSheet.Cells(3, 2).Value = MaxPercentIncrease & "%"
    
    resultSheet.Cells(1, 4).Value = "Greatest % Decrease"
    resultSheet.Cells(2, 4).Value = "Ticker"
    resultSheet.Cells(2, 5).Value = "Value"
    resultSheet.Cells(3, 4).Value = TickerMaxPercentDecrease
    resultSheet.Cells(3, 5).Value = MaxPercentDecrease & "%"
    
    resultSheet.Cells(1, 7).Value = "Greatest Total Volume"
    resultSheet.Cells(2, 7).Value = "Ticker"
    resultSheet.Cells(2, 8).Value = "Value"
    resultSheet.Cells(3, 7).Value = TickerMaxTotalVolume
    resultSheet.Cells(3, 8).Value = MaxTotalVolume

    ' Format the output for better visibility
    resultSheet.Columns("A:H").AutoFit
End Sub
