Sub analysis_on_multiple_sheets()

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim percentchange As Double
    Dim TotalStockVol As Double
    Dim increase As Double
    Dim decrease As Double
    Dim greatestvol As Double
    Dim firstValue As Double
    
    Dim SummaryTable As Long
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        ws.Activate
        
        SummaryTable = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
    
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                
                TotalStockVol = TotalStockVol + ws.Cells(i, 7)
                YearlyChange = (ws.Cells(i, 6).Value) - firstValue
                percentchange = ((ws.Cells(i, 6).Value) - firstValue) / firstValue
                
                ws.Range("I" & SummaryTable).Value = Ticker
                ws.Range("L" & SummaryTable).Value = TotalStockVol
                ws.Range("J" & SummaryTable).Value = YearlyChange
                ws.Range("K" & SummaryTable).Value = percentchange
                
                If percentchange > 0 Then
                
                    ws.Range("K" & SummaryTable).Interior.Color = RGB(0, 200, 0)
            
                ElseIf percentchange < 0 Then
                
                    ws.Range("K" & SummaryTable).Interior.Color = RGB(200, 0, 0)
            
                End If
                                   
                SummaryTable = SummaryTable + 1
                
                TotalStockVol = 0
                firstValue = ws.Cells(i + 1, 3).Value
                
            ElseIf i = 2 Then
                
                firstValue = ws.Cells(i, 3).Value
                
            Else
            
                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            End If
                     
        Next i
        
        increase = Application.WorksheetFunction.Max(Range("K:K"))
        decrease = Application.WorksheetFunction.Min(Range("K:K"))
        greatestvol = Application.WorksheetFunction.Max(Range("L:L"))
        
        
        For j = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
                
            If Cells(j, 11) = increase Then
            
                ws.Range("Q2") = increase
                ws.Range("P2") = ws.Cells(j, 9).Value
                
            ElseIf Cells(j, 11) = decrease Then
            
                ws.Range("Q3") = decrease
                ws.Range("P3") = ws.Cells(j, 9).Value
                
            ElseIf Cells(j, 12) = greatestvol Then
            
                ws.Range("Q4") = greatestvol
                ws.Range("P4") = ws.Cells(j, 9).Value
                
            End If
            
            Next j
        
        ws.Range("I:Q").EntireColumn.AutoFit
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws
    
End Sub