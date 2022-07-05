Sub Homework2()

For Each ws In Worksheets
         Dim worksheetCount As Integer
         Dim Ticker As Integer
         Dim YearlyChange As Double
         Dim PercentageChange As Double
         Dim OpeningPrice As Double
         Dim ClosingPrice As Double
         Dim StockVolume As Double
         Dim Counter As Integer
         Dim GreatIncr As Double
         Dim GreatDecr As Double
         Dim GreatVol As Double
           
           
    WorksheetName = ws.Name
          
          LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
                summary_table_row = 2
                ticker_row = 2
                YearlyChange = 0
                Percentage = 0
                OpeningPrice = 0
                numTicker = 0
                StockVolume = 0
                ws.Cells(1, "I").Value = "ticker"
                ws.Cells(1, "J").Value = "yearly Change"
                ws.Cells(1, "K").Value = "percent Change"
                ws.Cells(1, "L").Value = "Total Stock Volume"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                ws.Cells(2, 15).Value = "Greatest % Increase"
                ws.Cells(3, 15).Value = "Greatest % Decrease"
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                 
                 j = 2
                 
     For i = 2 To LastRowA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
            'stockvolume = (sum of stockvolume from i=2 to 252 from Else function) + cells(252, 7) = 765628638
                
        YearlyChange = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    
                        If ws.Cells(summary_table_row, 10).Value < 0 Then
                
                        'Set cell background color to red
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                
                        Else
                
                        'Set cell background color to green
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                
                        End If
                    'Calculate and write percent change in column K
                        If ws.Cells(j, 3).Value <> 0 Then
                        PercentageChange = (YearlyChange / ws.Cells(j, 3).Value)
                      
                    ws.Cells(summary_table_row, 11).Value = Format(PercentageChange, "Percent")
                        Else
                    
                        ws.Cells(summary_table_row, 11).Value = Format(0, "Percent")
                    
                        End If
                    '-----------------------------------------
                ws.Range("K" & summary_table_row).Value = PercentageChange
                ws.Range("J" & summary_table_row).Value = YearlyChange
                ws.Range("I" & summary_table_row).Value = ws.Cells(i, 1).Value
                ws.Range("L" & summary_table_row).Value = StockVolume
                    
                    
                summary_table_row = summary_table_row + 1
                StockVolume = 0

                j = i + 1
  Else
                    

     StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        End If

    Next i
    
   LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
       '
        
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                
                If ws.Cells(i, 12).Value > GreatVol Then
                
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
                    
     Next ws
     
End Sub
