Sub Stock_challenge()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
        Dim i As Long
        Dim ticker As String
        Dim ticker_position As Integer
        Dim ticker_open As Double
        Dim ticker_close As Double
        Dim ticker_volume As Double
        ticker_position = 2
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        ticker_volume = 0
        ticker = Cells(2, 1).Value
        ticker_open = Cells(2, 3).Value
        For i = 2 To lastrow
            ticker_volume = ticker_volume + Cells(i, 7).Value
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker_close = Cells(i, 6).Value
                Cells(ticker_position, 9).Value = ticker
                Cells(ticker_position, 10).Value = ticker_close - ticker_open
                If Cells(ticker_position, 10).Value > 0 Then
                    Cells(ticker_position, 10).Interior.ColorIndex = 4
                Else
                    Cells(ticker_position, 10).Interior.ColorIndex = 3
                End If
                If ticker_open <> 0 Then
                    Cells(ticker_position, 11).Value = ((ticker_close - ticker_open) / ticker_open) * 100
                End If
                Cells(ticker_position, 12).Value = ticker_volume
                ticker = Cells(i + 1, 1).Value
                ticker_open = Cells(i + 1, 3).Value
                ticker_volume = 0
                ticker_position = ticker_position + 1
            End If
        Next i
        
        
        Dim max As Double
        
        max = Cells(2, 11).Value
        Min = Cells(2, 11).Value
        max_volume = Cells(2, 12).Value
        Cells(2, 13).Value = "Greatest % Increase"
        Cells(3, 13).Value = "Greatest % Decrease"
        Cells(4, 13).Value = "Grearest Total Volume"
        For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row
           
            If Cells(i, 11).Value > max Then
                max = Cells(i, 11).Value
                ticker_max = Cells(i, 9).Value
            ElseIf Cells(i, 11) < Min Then
                Min = Cells(i, 11).Value
                ticker_min = Cells(i, 9).Value
            ElseIf Cells(i, 12) > max_volume Then
                max_volume = Cells(i, 12).Value
                ticker_greatestvolume = Cells(i, 9).Value
            End If
        Next i
        Cells(2, 14).Value = max
        Cells(3, 14).Value = Min
        Cells(4, 14).Value = max_volume
        
        Cells(2, 15).Value = ticker_max
        Cells(3, 15).Value = ticker_min
        Cells(4, 15).Value = ticker_greatestvolume
    Next
End Sub


