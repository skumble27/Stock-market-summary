Attribute VB_Name = "Module1"
Sub StockMarketSummary()

For Each ws In Worksheets


Dim ticker As String
Dim rowcount As Long
Dim lastrow As Long
Dim summary_table_row As Integer
Dim yearchange As Double
Dim percentchange As Double
Dim stockvolume As Double
Dim openprice As Double
Dim closedprice As Double
Dim summaryrowcount As Long





'The ticker count will assist in calculating the yearly change in stock price
rowcount = 0

'Summary Table Row Variable will fascilitate in entering new data in the next row
summary_table_row = 2

'We want the script to count the number of rows to end the for loop when required
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



'Initialise the stock volume as zero
stockvolume = 0


ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 10).Font.Bold = True

ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 11).Font.Bold = True

ws.Cells(1, 12).Value = "Percentage Change"
ws.Cells(1, 12).Font.Bold = True



ws.Cells(1, 13).Value = "Stock Volume"
ws.Cells(1, 13).Font.Bold = True

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Range("O:O").EntireColumn.AutoFit

For I = 2 To lastrow
    
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value And ws.Cells(I - rowcount, 3).Value = 0 Then
    
    ticker = ws.Cells(I, 1).Value
    stockvolume = stockvolume + (ws.Cells(I, 7).Value)
    
    
    
    ws.Range("J" & summary_table_row).Value = ticker
    ws.Range("K" & summary_table_row).Value = 0
    ws.Range("L" & summary_table_row).Value = 0
    ws.Range("M" & summary_table_row).Value = stockvolume
    ws.Range("N" & summary_table_row).Value = "No Open Price"
    
    
    summary_table_row = summary_table_row + 1
    rowcount = 0
    stockvolume = 0
    yearchange = 0
    
    
    
    
    ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
       
            
    ticker = ws.Cells(I, 1).Value
    openprice = ws.Cells(I - rowcount, 3).Value
    closedprice = ws.Cells(I, 6).Value
    yearchange = closedprice - openprice
    
    
       
    percentchange = (yearchange / openprice)
    stockvolume = stockvolume + (ws.Cells(I, 7).Value)
    
    
    
    
    
    ws.Range("J" & summary_table_row).Value = ticker
    ws.Range("K" & summary_table_row).Value = yearchange
    ws.Range("L" & summary_table_row).Value = percentchange
    ws.Range("M" & summary_table_row).Value = stockvolume
      
    
    
    summary_table_row = summary_table_row + 1
    rowcount = 0
    stockvolume = 0
    
    
       
        
    Else
    
    rowcount = rowcount + 1
    stockvolume = stockvolume + (ws.Cells(I, 7).Value)
    closingprice = closingprice + (ws.Cells(I, 6).Value)
    
    
    
    
    
    
    
    
    End If
    
    
Next I

summaryrowcount = ws.Cells(Rows.Count, 10).End(xlUp).Row

For I = 2 To summaryrowcount

    If ws.Cells(I, 12) > Max Then
    Max = ws.Cells(I, 12)
    ws.Cells(2, 16) = Max
    
    
    End If
       
Next I

For I = 2 To summaryrowcount
    
    If ws.Cells(I, 12) < Min Then
    Min = ws.Cells(I, 12)
    ws.Cells(3, 16) = Min
    
    End If
Next I

For I = 2 To summaryrowcount

    If ws.Cells(I, 13) > Max Then
    Max = ws.Cells(I, 13)
    ws.Cells(4, 16) = Max
    
    End If
Next I


ws.Range("J:M").EntireColumn.AutoFit
ws.Range("k2", ws.Range("K2").End(xlDown)).NumberFormat = "0.00"
ws.Range("L2", ws.Range("L2").End(xlDown)).NumberFormat = "0.00%"
ws.Range("M2", ws.Range("M2").End(xlDown)).NumberFormat = "0.00E+00"
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
ws.Range("P4").NumberFormat = "0.00E+00"

For I = 2 To summaryrowcount

    If ws.Cells(I, 11).Value < 0 Then
    ws.Cells(I, 11).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(I, 11).Value > 0 Then
    ws.Cells(I, 11).Interior.ColorIndex = 4
    
       
    ElseIf ws.Cells(I, 11).Value = 0 Then
    ws.Cells(I, 10).Interior.ColorIndex = 34
    ws.Cells(I, 11).Interior.ColorIndex = 34
    ws.Cells(I, 12).Interior.ColorIndex = 34
    ws.Cells(I, 13).Interior.ColorIndex = 34
    
    
       
    End If
Next I

For I = 2 To summaryrowcount

    If ws.Cells(I, 14).Value = "No Open Price" Then
    ws.Cells(I, 10).Interior.ColorIndex = 24
    ws.Cells(I, 11).Interior.ColorIndex = 24
    ws.Cells(I, 12).Interior.ColorIndex = 24
    ws.Cells(I, 13).Interior.ColorIndex = 24
    ws.Cells(I, 14).Interior.ColorIndex = 24
    ws.Range("N:N").EntireColumn.AutoFit
    
    
    End If
Next I

Max = 0
Min = 0

MsgBox ("Summary for Sheet " & ws.Name & " Completed")



Next ws

End Sub




