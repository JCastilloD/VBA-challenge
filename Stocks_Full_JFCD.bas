Attribute VB_Name = "Module1"
Sub Stocks()
    Dim i, ii, Total
    Dim Max(2 To 4)
    Dim Ticker(2 To 4) As String
    Dim ws As Worksheet

    ii = 2
    
    For Each ws In Worksheets
        i = 2
        ii = 2
        ws.Cells(1, 10) = "Ticker"
        ws.Cells(1, 11) = "Yearly Change"
        ws.Cells(1, 12) = "Percent change"
        ws.Cells(1, 13) = "Total Stock Volume"
        
        ws.Cells(ii, 10) = ws.Cells(i, 1)
        ws.Cells(ii, 11) = ws.Cells(i, 3)
        Total = ws.Cells(i, 7)
        For i = 3 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            Total = Total + ws.Cells(i, 7)
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                ws.Cells(ii + 1, 10) = ws.Cells(i + 1, 1)
                ws.Cells(ii + 1, 11) = ws.Cells(i + 1, 3)
                If ws.Cells(ii, 11) <> 0 Then
                    ws.Cells(ii, 12) = (ws.Cells(i, 6) / ws.Cells(ii, 11)) - 1
                Else
                    'Cells(ii, 12) = (ws.Cells(i, 6) / 0.01) - 1
                    'Cells(ii, 12) = "New" ' Distintas Respuestas según se
                    'Cells(ii, 12) = "N/A" ' requiera para compañias nuevas
                    ws.Cells(ii, 12) = 0
                End If
                ws.Cells(ii, 12).NumberFormat = "0.00%"
                
                If ws.Cells(ii, 12) > Max(2) Then
                    Max(2) = ws.Cells(ii, 12)
                    Ticker(2) = ws.Cells(ii, 10)
                ElseIf ws.Cells(ii, 12) < Max(3) Then
                    Max(3) = ws.Cells(ii, 12)
                    Ticker(3) = ws.Cells(ii, 10)
                End If
                
                ws.Cells(ii, 11) = ws.Cells(i, 6) - ws.Cells(ii, 11)
                If ws.Cells(ii, 11) > 0 Then
                    ws.Cells(ii, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf ws.Cells(ii, 11) < 0 Then
                    ws.Cells(ii, 11).Interior.Color = RGB(255, 0, 0)
                Else
                    ' Si el cambio es 0 ss pone Amarillo:
                    ws.Cells(ii, 11).Interior.Color = RGB(255, 255, 0)
                End If
                    
                If Max(4) < Total Then
                    Max(4) = Total
                    Ticker(4) = ws.Cells(ii, 10)
                End If
                
                ws.Cells(ii, 13) = Total
                Total = ws.Cells(i + 1, 7)
                ii = ii + 1
                
            End If
        Next i
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("M:M").EntireColumn.AutoFit
        
        For i = 2 To 4
            ws.Cells(i, 17) = Ticker(i)
            ws.Cells(i, 18) = Max(i)
            Max(i) = 0
        Next i
        
        ws.Cells(1, 17) = "Ticker"
        ws.Cells(1, 18) = "Value"
        ws.Cells(2, 16) = "Greatest % increase"
        ws.Cells(3, 16) = "Greatest % Decrease"
        ws.Cells(4, 16) = "Greatest Total Volume"
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 18).NumberFormat = "0.00E+00"
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        ws.Columns("R:R").EntireColumn.AutoFit
        
    Next ws
    
End Sub
