Attribute VB_Name = "Final"
Sub StockMarketAnalysis():


        Dim GIncrease As Double
        GIncrease = Cells(2, 11).Value
        
        Dim GDecrease As Double
        GDecrease = Cells(2, 11).Value
        
        Dim GTotalVolume As Double
        GTotalVolume = Cells(2, 12).Value
        
    
    
    For Each ws In Worksheets
    
    
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        
        Dim i As Long
        
        Dim CounterTicker As Integer
        CounterTicker = 2
        
        Dim CounterVolume As Integer
        CounterVolume = 2
        
        Dim CounterYearly As Integer
        CounterYearly = 2
        
        Dim CounterPercentage As Integer
        CounterPercentage = 2
        
        Dim Volume As Double
        Volume = 0
        
        Dim stockOpen As Double
        stockOpen = 0
        
        Dim stockClose As Double
        stockClose = 0
        
        Dim YearlyChange As Double
        YearlyChange = 0
        
        Dim percentageOpen As Double
        percentageOpen = 0
        
        Dim percentageClose As Double
        percentageClose = 0
        
        For i = 2 To LastRow
            If IsEmpty(ws.Range("I1").Value) Then
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
        End If
        
            If ws.Cells(i, 2).Value = "20180102" Then
                ws.Cells(CounterTicker, 9).Value = ws.Cells(i, 1).Value
                CounterTicker = CounterTicker + 1
            End If
            
             If ws.Cells(i, 2).Value = "20190102" Then
                ws.Cells(CounterTicker, 9).Value = ws.Cells(i, 1).Value
                CounterTicker = CounterTicker + 1
            End If
            
             If ws.Cells(i, 2).Value = "20200102" Then
                ws.Cells(CounterTicker, 9).Value = ws.Cells(i, 1).Value
                CounterTicker = CounterTicker + 1
            End If
            
            
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            Volume = Volume + ws.Cells(i, 7).Value
            Else
                Volume = Volume + ws.Cells(i, 7).Value
                ws.Cells(CounterVolume, 12).Value = Format(Volume, "#,##0")
                CounterVolume = CounterVolume + 1
                Volume = 0
            End If
            
            
            
            
            If ws.Cells(i, 2).Value = "20200102" Then
                stockOpen = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = "20201231" Then
                stockClose = ws.Cells(i, 6).Value
                YearlyChange = stockClose - stockOpen
                ws.Cells(CounterYearly, 10).Value = Format(YearlyChange, "#,#0")
                If ws.Cells(CounterYearly, 10).Value < 0 Then
                    ws.Cells(CounterYearly, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(CounterYearly, 10).Interior.ColorIndex = 4
                End If
                
          
             stockOpen = 0
             stockClose = 0
             YearlyChange = 0
             CounterYearly = CounterYearly + 1
                End If
                
                
                
              If ws.Cells(i, 2).Value = "20190102" Then
                stockOpen = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = "20191231" Then
                stockClose = ws.Cells(i, 6).Value
                YearlyChange = stockClose - stockOpen
                ws.Cells(CounterYearly, 10).Value = Format(YearlyChange, "#,#0")
                If ws.Cells(CounterYearly, 10).Value < 0 Then
                    ws.Cells(CounterYearly, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(CounterYearly, 10).Interior.ColorIndex = 4
                End If
                
             stockOpen = 0
             stockClose = 0
             YearlyChange = 0
             CounterYearly = CounterYearly + 1
                End If
                
                
            If ws.Cells(i, 2).Value = "20180102" Then
                percentageOpen = ws.Cells(i, 3).Value
                
            ElseIf ws.Cells(i, 2).Value = "20181231" Then
                    percentageClose = ws.Cells(i, 6).Value
                    percentageChange = ((percentageClose - percentageOpen) / percentageOpen) * 100
                    ws.Cells(CounterPercentage, 11).Value = Format(percentageChange, "0,00") & "%"
                    If ws.Cells(CounterPercentage, 11) < 0 Then
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 4
                End If
                
                percentageOpen = 0
                percentageClose = 0
                percentageChange = 0
                CounterPercentage = CounterPercentage + 1
                End If
                
            
                If ws.Cells(i, 2).Value = "20190102" Then
                percentageOpen = ws.Cells(i, 3).Value
                
            ElseIf ws.Cells(i, 2).Value = "20191231" Then
                    percentageClose = ws.Cells(i, 6).Value
                    percentageChange = ((percentageClose - percentageOpen) / percentageOpen) * 100
                    ws.Cells(CounterPercentage, 11).Value = Format(percentageChange, "0,00") & "%"
                    If ws.Cells(CounterPercentage, 11) < 0 Then
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 4
                End If
                
                percentageOpen = 0
                percentageClose = 0
                percentageChange = 0
                CounterPercentage = CounterPercentage + 1
                End If
                
                    If ws.Cells(i, 2).Value = "20200102" Then
                percentageOpen = ws.Cells(i, 3).Value
                
            ElseIf ws.Cells(i, 2).Value = "20201231" Then
                    percentageClose = ws.Cells(i, 6).Value
                    percentageChange = ((percentageClose - percentageOpen) / percentageOpen) * 100
                    ws.Cells(CounterPercentage, 11).Value = Format(percentageChange, "0,00") & "%"
                    If ws.Cells(CounterPercentage, 11) < 0 Then
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(CounterPercentage, 11).Interior.ColorIndex = 4
                End If
                
                percentageOpen = 0
                percentageClose = 0
                percentageChange = 0
                CounterPercentage = CounterPercentage + 1
                End If
    
               
             Next i
             
             
             
             For j = 3 To LastRow
                If ws.Cells(j, 11).Value > GIncrease Then
                GIncrease = ws.Cells(j, 11).Value
                Cells(2, 17).Value = Format(GIncrease * 100, "0.00" & "%")
                Cells(2, 16).Value = ws.Cells(j, 9).Value
                End If
                
                If ws.Cells(j, 11).Value < GDecrease Then
                GDecrease = ws.Cells(j, 11).Value
                Cells(3, 17).Value = Format(GDecrease * 100, "0.00" & "%")
                Cells(3, 16).Value = ws.Cells(j, 9).Value
                End If
                
                If ws.Cells(j, 12).Value > GTotalVolume Then
                GTotalVolume = ws.Cells(j, 12).Value
                Cells(4, 17).Value = Format(GTotalVolume * 100, "0.00" & "%")
                Cells(4, 16).Value = ws.Cells(j, 9).Value
                End If
                
        Next j
        
    Next ws
    
                      
  End Sub
  
  
            
            
        
        
        
        
        
        
        
        
        
        
        




