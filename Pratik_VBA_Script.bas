Attribute VB_Name = "Module1"
Sub VBA_challenge_2()

For Each ws In Worksheets

    'labels for headers
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total stock volume"

    'declaring variables
    Dim TickerName As String
    Dim LastRow As Long
    Dim totalticker As Double
    totalticker = 0
    Dim summaryrow As Long
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim PreviousAmount As Long
    Dim PercentChange As Double
    PreviousAmount = 2
    summaryrow = 2

    'finding finalrow

    FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To FinalRow

    totalticker = totalticker + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


     ' tickername
     TickerName = ws.Cells(i, 1).Value
     ws.Range("I" & summaryrow).Value = TickerName
     ws.Range("L" & summaryrow).Value = totalticker
     totalticker = 0

    YearlyOpen = ws.Range("C" & PreviousAmount)
    YearlyClose = ws.Range("F" & i)
    YearlyChange = YearlyClose - YearlyOpen
    ws.Range("J" & summaryrow).Value = YearlyChange

    If YearlyChange = 0 Then
        PercentChange = 0

        Else
         YearlyOpen = ws.Range("C" & PreviousAmount)
         PercentChange = YearlyChange / YearlyOpen

         End If
       ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("K" & summaryrow).NumberFormat = "0.00%"
            ws.Range("K" & summaryrow).Value = PercentChange

                ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
        If ws.Range("J" & summaryrow).Value >= 0 Then
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 4

                     Else
                         ws.Range("J" & summaryrow).Interior.ColorIndex = 3

                        End If

                ' Add One To The Summary Table Row
                summaryrow = summaryrow + 1
                PreviousAmount = i + 1
                End If
            Next i

        Next ws




End Sub
