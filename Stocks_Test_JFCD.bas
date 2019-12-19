Attribute VB_Name = "Module1"
Sub Stocks()
    Dim i, ii, Total
    Dim Max(2 To 4)
    Dim Ticker(2 To 4) As String
    Dim ws As Worksheet

    ii = 2
    
    Cells(1, 17) = "Ticker"
    Cells(1, 18) = "Value"
    Cells(2, 16) = "Greatest % increase"
    Cells(3, 16) = "Greatest % Decrease"
    Cells(4, 16) = "Greatest Total Volume"
    Columns("P:P").EntireColumn.AutoFit
    For Each ws In Worksheets
        i = 2
        Cells(1, 10) = "Ticker"
        Cells(1, 11) = "Yearly Change"
        Cells(1, 12) = "Percent change"
        Cells(1, 13) = "Total Stock Volume"
        
        Cells(ii, 10) = ws.Cells(i, 1)
        Cells(ii, 11) = ws.Cells(i, 3)
        Total = ws.Cells(i, 7)
        For i = 3 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            Total = Total + ws.Cells(i, 7)
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                Cells(ii + 1, 10) = ws.Cells(i + 1, 1)
                Cells(ii + 1, 11) = ws.Cells(i + 1, 3)
                If Cells(ii, 11) <> 0 Then
                    Cells(ii, 12) = (ws.Cells(i, 6) / Cells(ii, 11)) - 1
                Else
                    'Cells(ii, 12) = (ws.Cells(i, 6) / 0.01) - 1
                    'Cells(ii, 12) = "New" ' Distintas Respuestas según se
                    'Cells(ii, 12) = "N/A" ' requiera para compañias nuevas
                    Cells(ii, 12) = 0
                End If
                Cells(ii, 12).NumberFormat = "0.00%"
                
                If Cells(ii, 12) > Max(2) Then
                    Max(2) = Cells(ii, 12)
                    Ticker(2) = Cells(ii, 10)
                ElseIf Cells(ii, 12) < Max(3) Then
                    Max(3) = Cells(ii, 12)
                    Ticker(3) = Cells(ii, 10)
                End If
                
                Cells(ii, 11) = ws.Cells(i, 6) - Cells(ii, 11)
                If Cells(ii, 11) > 0 Then
                    Cells(ii, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf Cells(ii, 11) < 0 Then
                    Cells(ii, 11).Interior.Color = RGB(255, 0, 0)
                Else
                    ' Si el cambio es 0 ss pone Amarillo:
                    Cells(ii, 11).Interior.Color = RGB(255, 255, 0)
                End If
                    
                If Max(4) < Total Then
                    Max(4) = Total
                    Ticker(4) = Cells(ii, 10)
                End If
                
                Cells(ii, 13) = Total
                Total = ws.Cells(i + 1, 7)
                ii = ii + 1
                
            End If
        Next i
        Columns("K:K").EntireColumn.AutoFit
        Columns("L:L").EntireColumn.AutoFit
        Columns("M:M").EntireColumn.AutoFit
    Next ws
    For i = 2 To 4
        Cells(i, 17) = Ticker(i)
        Cells(i, 18) = Max(i)
    Next i
    Cells(2, 18).NumberFormat = "0.00%"
    Cells(3, 18).NumberFormat = "0.00%"
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
End Sub
