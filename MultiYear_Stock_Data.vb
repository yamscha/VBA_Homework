Sub Multiyear_Stock_MarketData()

     Dim TickerName As String
     Dim TickerTotal As Double
     Dim Worksheetname As String
     Dim lastRow1 As Long

For Each ws In Worksheets
    With ws
    
    LastRow = .Cells(Rows.Count, 1).End(xlDown).Row
        
    TickerTotal = 0
    SumRow = 2
        
    lastRow1 = .Range("A2").End(xlDown).Row

    TickerName = .Cells(2, 1).Value
    
    .Range("I1").Value = "Ticker"
    .Range("J1").Value = "TotalStockVolume"
    
    For i = 2 To lastRow1
      If TickerName = .Cells(i + 1, 1).Value Then
         TickerTotal = TickerTotal + .Cells(i, 7).Value
            
      Else
          TickerTotal = TickerTotal + ws.Cells(i, 7).Value
          .Range("I2").Cells(SumRow, 1).Value = TickerName
          .Range("I2").Cells(SumRow, 2).Value = TickerTotal
            SumRow = SumRow + 1
            TickerTotal = 0
            
            TickerName = .Cells(i + 1, 1).Value
       End If
      Next i
      End With
Next ws
 
 End Sub
 

 
