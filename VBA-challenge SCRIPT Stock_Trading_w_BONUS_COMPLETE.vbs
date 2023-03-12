Sub Stock_Trading_w_BONUS_COMPLETE()

'-----Loop through all sheets
  Dim ws As Worksheet
  For Each ws In Worksheets
    
    '-----Create variables for Ticker, OpenPrice, ClosePrice and StockVolume
    Dim Ticker As String
    Dim i As Long
    Dim LastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    'BONUS
    Dim PercentI As Double
    Dim PercentD As Double
    Dim TotalVol As Double
        
     '-----PRINT Table headers: Ticker, Yearly Change, Percent Change and Total Stock Volume
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     ws.Cells(2, 14) = "Greatest Percent Increase"     'BONUS - Table header
     ws.Cells(3, 14) = "Greatest Percent Decrease"     'BONUS - Table header
     ws.Cells(4, 14) = "Greatest Total Volume"         'BONUS - Table header
     ws.Cells(1, 15) = "Ticker"                        'BONUS - Table header
     ws.Cells(1, 16) = "Value"                         'BONUS - Table header
       
       '-----Determine Last Row
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
         '-----SET Values & location for Table
         UniqueTickerRow = 2
         StockVolume = 0
         YearChange = 0
         PercentChange = 0
         'BONUS
         PercentI = 0
         PercentD = 0
         TotalVol = 0
        
          '-----Loop through rows
          For i = 2 To LastRow
            '-----SET values for 1st data row
            If OpenPrice = 0 Then
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            StockVolume = StockVolume + Cells(i, 7).Value
            YearChange = ClosePrice - OpenPrice
            PercentChange = YearChange / OpenPrice
            
             '--------------------------------
             'Next Ticker code in loop UNIQUE
             '--------------------------------
              ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                '-----SET UniqueTicker & HOLD values
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearChange = ClosePrice - OpenPrice
                PercentChange = YearChange / OpenPrice
                StockVolume = StockVolume + Cells(i, 7).Value
                   '-----PRINT values into UniqueTicker Table & only display 2 decimal places
                  ws.Range("I" & UniqueTickerRow).Value = Ticker
                  ws.Range("L" & UniqueTickerRow).Value = StockVolume
                  ws.Range("J" & UniqueTickerRow).Value = YearChange
                  ws.Range("K" & UniqueTickerRow).Value = FormatPercent(PercentChange)
                'BONUS TABLE SET & PRINT Ticker and Hold Values
                   If PercentChange > ws.Range("P2").Value Then  '-----Greatest Percent Increase
                       PercentI = PercentChange
                        ws.Range("O2").Value = Ticker
                        ws.Range("P2").Value = FormatPercent(PercentI)
                         End If
                   If PercentChange < ws.Range("P3").Value Then  '-----Greatest Percent Decrease
                       PercentD = PercentChange
                        ws.Range("O3").Value = Ticker
                        ws.Range("P3").Value = FormatPercent(PercentD)
                         End If
                   If StockVolume > ws.Range("P4").Value Then  '-----Greatest Total Volume
                       TotalVol = StockVolume
                        ws.Range("O4").Value = Ticker
                        ws.Range("P4").Value = TotalVol
                         End If
                     '-----HIGHLIGHT positive change: green / negative change: red / no change: white
                     If YearChange < 0 Then
                        ws.Range("J" & UniqueTickerRow).Interior.ColorIndex = 3
                      ElseIf YearChange > 0 Then
                        ws.Range("J" & UniqueTickerRow).Interior.ColorIndex = 35
                       Else
                        ws.Range("J" & UniqueTickerRow).Interior.ColorIndex = 2
                     End If
                        '-----SET next row & RESET values
                        UniqueTickerRow = UniqueTickerRow + 1
                        StockVolume = 0
                        OpenPrice = 0
                        ClosePrice = 0
                        YearChange = 0
                        PercentChange = 0

                 '------------------------------
                 'Next Ticker code in loop SAME
                 '------------------------------
                 Else
                    'UPDATE held values (EXCLUDING OPEN PRICE)
                    StockVolume = StockVolume + ws.Cells(i, 7).Value
                    ClosePrice = ws.Cells(i, 6).Value
                    YearChange = ClosePrice - OpenPrice
                    PercentChange = YearChange / OpenPrice
            End If
                                         
          Next i
                
            '-----FORMATING
            ws.Range("A:P").Columns.AutoFit 'Autofit column width
      '-----
    Next ws
   
End Sub


