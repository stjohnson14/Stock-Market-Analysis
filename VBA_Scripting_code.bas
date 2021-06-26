Attribute VB_Name = "Module1"
Sub Stock_Counter()

For Each ws In Worksheets

Dim ticker As String
Dim year_change As Variant
Dim year_open As Double
Dim year_close As Double
Dim percent_change As Variant
Dim lastrow As Long
Dim tickercount As Integer
Dim total_volume As Variant

WorksheetName = ws.Name

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

tickercount = 1
summaryrowtable = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("I1").Font.Bold = True
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J1").Font.Bold = True
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K1").Font.Bold = True
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L1").Font.Bold = True
    
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

   total_volume = total_volume + ws.Cells(i, 7).Value
    
    year_open = year_open + ws.Cells(i, 3).Value
    year_close = year_close + ws.Cells(i, 6).Value
    
    year_change = ((year_close - year_open) / tickercount) * 100
    year_change = Format(year_change, "#0.00")
    
    On Error Resume Next
    percent_change = ((year_close - year_open) / year_open)
    percent_change = percent_change * 100
    percent_change = Format(percent_change, "#0.00")
    
    ws.Range("J" & summaryrowtable) = year_change
    ws.Range("K" & summaryrowtable) = percent_change
    ws.Range("K" & summaryrowtable).NumberFormat = "0.00%"
    ws.Range("I" & summaryrowtable) = ws.Cells(i, 1).Value
    ws.Range("L" & summaryrowtable) = total_volume
    
    ws.Cells(i, 10).Value = total_volume

    summaryrowtable = summaryrowtable + 1
    
    total_volume = 0
    year_open = 0
    year_close = 0
    year_change = 0
    tickercount = 0

Else

tickercount = tickercount + 1
year_open = year_open + ws.Cells(i, 3).Value
year_close = year_close + ws.Cells(i, 6).Value
total_volume = total_volume + ws.Cells(i, 7).Value


End If

If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i
 
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O2").Font.Bold = True
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O3").Font.Bold = True
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("O4").Font.Bold = True
    ws.Range("P1").Value = "Ticker"
    ws.Range("P1").Font.Bold = True
    ws.Range("Q1").Value = "Value"
    ws.Range("Q1").Font.Bold = True

    lastrowbonus = ws.Cells(Rows.Count, "I").End(xlUp).Row
   
    greatest_percent_increase = ws.Cells(2, 11).Value
    greatest_percent_increase_ticker = ws.Cells(2, 9).Value
    greatest_percent_decrease = ws.Cells(2, 11).Value
    greatest_percent_decrease_ticker = ws.Cells(2, 9).Value
    greatest_stock_volume = ws.Cells(2, 12).Value
    greatest_stock_volume_ticker = ws.Cells(2, 9).Value
    
For i = 2 To lastrowbonus
        
    If ws.Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = ws.Cells(i, 11).Value
            greatest_percent_increase_ticker = ws.Cells(i, 9).Value
    End If
        
    If ws.Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = ws.Cells(i, 11).Value
            greatest_percent_decrease_ticker = ws.Cells(i, 9).Value
     End If
    
    If ws.Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = ws.Cells(i, 12).Value
            greatest_stock_volume_ticker = ws.Cells(i, 9).Value
    End If
        
Next i
  
    ws.Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    ws.Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    ws.Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    ws.Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    ws.Range("P4").Value = greatest_stock_volume_ticker
    ws.Range("Q4").Value = greatest_stock_volume

Next ws

End Sub



