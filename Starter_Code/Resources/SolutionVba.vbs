Attribute VB_Name = "Module1"
Sub Lezduit()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Yearchange As Double
    Dim Perchange As Double
    Dim Stock As LongLong
    Dim table As Integer
    Dim start As Double
    Dim final As Double

    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

        ' Titles of table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        table = 2

        last1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        For i = 2 To last1
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Stock = Stock + ws.Cells(i, 7).Value
                ws.Range("I" & table).Value = Ticker
                ws.Range("L" & table).Value = Stock
                final = ws.Cells(i, 6).Value
                Perchange = final / start - 1
                Yearchange = final - start
                ws.Range("K" & table).Value = Perchange
                ws.Range("K" & table).NumberFormat = "0.00%"
                ws.Range("J" & table).Value = Yearchange
                ws.Range("J" & table).NumberFormat = "0.00"
                Yearchange = 0
                Perchange = 0
                table = table + 1
                Stock = 0
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                start = ws.Cells(i, 3).Value
            Else
                Stock = Stock + ws.Cells(i, 7).Value
            End If
        Next i

        table = 2

        last2 = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row

        For i = 2 To last2
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        Dim GperIn As Double
        Dim GperDe As Double
        Dim HighVo As LongLong
        Dim TickerI As String
        Dim TickerD As String
        Dim TickerV As String

        last2 = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        percentchange = ws.Range("K1:K" & last2)
        totalvolume = ws.Range("L1:L" & last2)

        GperIn = WorksheetFunction.Max(percentchange)
        GperDe = WorksheetFunction.Min(percentchange)
        HighVo = WorksheetFunction.Max(totalvolume)

        For i = 2 To last2
            If HighVo = ws.Cells(i, 12) Then
                TickerV = ws.Cells(i, 9).Value
            End If

            For k = 1 To 2
                If GperIn = ws.Cells(i, 11) Then
                    TickerI = ws.Cells(i, 9).Value
                ElseIf GperDe = ws.Cells(i, 11) Then
                    TickerD = ws.Cells(i, 9).Value
                End If
            Next k
        Next i

        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("P2").Value = TickerI
        ws.Range("P3").Value = TickerD
        ws.Range("Q2").Value = GperIn
        ws.Range("Q3").Value = GperDe
        ws.Range("P4").Value = TickerV
        ws.Range("Q4").Value = HighVo

        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        ws.Columns("I:I").EntireColumn.AutoFit
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
    Next ws

    MsgBox ("The program has finished")
End Sub
