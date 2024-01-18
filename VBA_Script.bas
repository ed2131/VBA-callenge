Attribute VB_Name = "Module1"
Sub Module1_VBA_Final_Solution()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim ticker As String, openingPrice As Double, closingPrice As Double, totalVolume As Double
    Dim maxPercentageIncrease As Double, maxPercentageDecrease As Double, maxTotalVolume As Double
    Dim tickerMaxIncrease As String, tickerMaxDecrease As String, tickerMaxVolume As String

    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Set Headers
            .Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            .Range("O1:Q1").Value = Array("Category", "Ticker", "Value")
            .Range("O2:O4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))

            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            outputRow = 2
            totalVolume = 0

            For i = 2 To lastRow
                ticker = .Cells(i, 1).Value
                If .Cells(i - 1, 1).Value <> ticker Then openingPrice = .Cells(i, 3).Value
                closingPrice = .Cells(i, 6).Value
                totalVolume = totalVolume + .Cells(i, 7).Value

                If .Cells(i + 1, 1).Value <> ticker Then
                    .Cells(outputRow, 9).Resize(1, 4).Value = Array(ticker, closingPrice - openingPrice, (closingPrice - openingPrice) / openingPrice, totalVolume)
                    .Cells(outputRow, 11).NumberFormat = "0.00%"
                    totalVolume = 0
                    outputRow = outputRow + 1
                End If
            Next i

            ' Conditional Formatting & Finding Max/Min Values
            maxPercentageIncrease = 0: maxPercentageDecrease = 0: maxTotalVolume = 0

            For i = 2 To outputRow - 1
                With .Cells(i, 10)
                    .Interior.Color = IIf(.Value < 0, RGB(255, 0, 0), RGB(0, 255, 0))
                End With
                With .Cells(i, 11)
                    .Interior.Color = IIf(.Value < 0, RGB(255, 0, 0), RGB(0, 255, 0))
                    If .Value > maxPercentageIncrease Then
                        maxPercentageIncrease = .Value
                        tickerMaxIncrease = .Offset(0, -2).Value
                    ElseIf .Value < maxPercentageDecrease Then
                        maxPercentageDecrease = .Value
                        tickerMaxDecrease = .Offset(0, -2).Value
                    End If
                End With
                If .Cells(i, 12).Value > maxTotalVolume Then
                    maxTotalVolume = .Cells(i, 12).Value
                    tickerMaxVolume = .Cells(i, 9).Value
                End If
            Next i

            .Range("P2").Resize(3, 1).Value = Application.Transpose(Array(tickerMaxIncrease, tickerMaxDecrease, tickerMaxVolume))
            .Range("Q2").Resize(3, 1).Value = Application.Transpose(Array(maxPercentageIncrease, maxPercentageDecrease, maxTotalVolume))
            .Range("Q2:Q3").NumberFormat = "0.00%"
            .Columns("A:Q").AutoFit
        End With
    Next ws

    MsgBox "Tabulation Completed"

End Sub

