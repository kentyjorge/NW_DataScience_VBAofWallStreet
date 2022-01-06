Attribute VB_Name = "stockLoop"
Sub stockLoop()

    Dim i, top, bottom, lastrow, tickrow As Long
    Dim tchange, tpchange, op, cl, tickertotal, maxt As Double
    Dim ticker As String

'insert worksheet loop
For Each ws In Worksheets

        'label headers for inputs
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
     
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tickrow = 2
        tickertotal = 0
        ws.Cells(2, 17).Value = 0
        ws.Cells(3, 17).Value = 0
        ws.Cells(4, 17).Value = 0
    
        For i = 2 To lastrow
            
            'sum total until the condition below is met, then set total back to 0
            tickertotal = tickertotal + ws.Cells(i, 7).Value
            
            'unique ticker symbols
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'setting ticker total
                ticker = Cells(i, 1).Value
                ws.Cells(tickrow, 9).Value = ticker
                ws.Cells(tickrow, 12).Value = tickertotal
                
                'setting greatest total
                If tickertotal > ws.Cells(4, 17).Value Then
                    ws.Cells(4, 17).Value = tickertotal
                    ws.Cells(4, 16).Value = ticker
                End If
    
                'find open and closed and then calc the change
                top = ws.Range("A:A").Find(What:=ws.Cells(i, 1).Value).Row  'first row of specific ticker
                op = ws.Cells(top, 3).Value
                cl = ws.Cells(i, 6).Value
                
                tchange = cl - op
                If op <> 0 Then
                tpchange = tchange / op
                Else
                    tpchange = 0
                End If

                ws.Cells(tickrow, 10).Value = tchange
                
               'colorcell
               If tchange > 0 Then
                    ws.Cells(tickrow, 10).Interior.ColorIndex = 4
               Else
                    ws.Cells(tickrow, 10).Interior.ColorIndex = 3
               End If
        
                ws.Cells(tickrow, 11).Value = tpchange
                
                'setting greatest increase and decrease % change
                If tpchange > Cells(2, 17).Value Then
                    ws.Cells(2, 17).Value = tpchange
                    ws.Cells(2, 16).Value = ticker
                ElseIf tpchange < Cells(3, 17).Value Then
                    ws.Cells(3, 17).Value = tpchange
                    ws.Cells(3, 16).Value = ticker
                End If
                
                'update counters
                tickrow = tickrow + 1
                tickertotal = 0
                
            End If
            
        Next i
    Next ws
    
End Sub
