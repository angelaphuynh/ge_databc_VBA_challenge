Attribute VB_Name = "Module1"
Sub stocks()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        Dim Open_Price, Close_Price, Yearly_Change, Percent_Change, Volume, Row As Double
        Dim Ticker_Name As String
        Dim i As Long
        Volume = 0
        Row = 2
        
        LRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Open_Price = Cells(2, 3).Value
        
        For i = 2 To LRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker
                Close_Price = Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, 10).Value = Yearly_Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                Cells(Row, 12).NumberFormat = "000,000"
                Row = Row + 1
                Open_Price = Cells(i + 1, 3)
                Volume = 0
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
            YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
            For j = 2 To YCLastRow
                If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                    Cells(j, 10).Interior.ColorIndex = 10
                ElseIf Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                End If
            Next j
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Columns("O:Q").EntireColumn.AutoFit
        
        For Z = 2 To YCLastRow
            If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Z, 9).Value
                Cells(2, 17).Value = Cells(Z, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Z, 9).Value
                Cells(3, 17).Value = Cells(Z, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Z, 9).Value
                Cells(4, 17).Value = Cells(Z, 12).Value
                 Cells(4, 17).NumberFormat = "000,000"
            End If
        Next Z
        
    Next WS
End Sub
