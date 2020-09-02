Attribute VB_Name = "Module1"
Sub StockMarketSummary()
'This macro will run through all the sheets in the workbook
For Each ws In Worksheets


Dim ticker As String 'will list the ticker and its associated summary data
Dim rowcount As Long 'This variable will fascilitate calculating the percentchange
Dim lastrow As Long 'To automatically count the number of rows
Dim summary_table_row As Integer 'To allow new data to be entered into the next row
Dim yearchange As Double
Dim percentchange As Double
Dim stockvolume As Double
Dim openprice As Double
Dim closedprice As Double
Dim summaryrowcount As Long ' To count the number of rows after all stocks have been summarised





rowcount = 0

summary_table_row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Initialise the stock volume as zero
stockvolume = 0

'Cells are being formatted and table titles are being added
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 10).Font.Bold = True

ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 11).Font.Bold = True

ws.Cells(1, 12).Value = "Percentage Change"
ws.Cells(1, 12).Font.Bold = True



ws.Cells(1, 13).Value = "Stock Volume"
ws.Cells(1, 13).Font.Bold = True

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 16).Font.Bold = True
ws.Cells(1, 17).Value = "Value"
ws.Cells(1, 17).Font.Bold = True
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Range("O:O").EntireColumn.AutoFit

For I = 2 To lastrow
    'In finance, percentchange change over the year requires an open price value.
        'This loop is to prevent the error that will occur when dividing by zero
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
    
    
    
    'This function will commence once the iterations move to a new stock
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
    
    
       
     'If the stock is still the same, this function will continue to run
    Else
    
    rowcount = rowcount + 1
    stockvolume = stockvolume + (ws.Cells(I, 7).Value)
    closingprice = closingprice + (ws.Cells(I, 6).Value)
    
    
    End If
    
    
Next I

'Once the loop has completed, the summary table rows can then be counted
summaryrowcount = ws.Cells(Rows.Count, 10).End(xlUp).Row

'This function determines the highest percentchange
For I = 2 To summaryrowcount

    If ws.Cells(I, 12) > Max Then
    Max = ws.Cells(I, 12)
    ws.Cells(2, 17) = Max
    ws.Cells(2, 16) = ws.Cells(I, 10).Value
    
    
    
    
    End If


Next I

'This function determines the lowest percentchange
For I = 2 To summaryrowcount
    
    If ws.Cells(I, 12) < Min Then
    Min = ws.Cells(I, 12)
    ws.Cells(3, 17) = Min
    ws.Cells(3, 16) = ws.Cells(I, 10).Value
    
    
    End If
Next I

'This function determines the highest stock volume
For I = 2 To summaryrowcount

    If ws.Cells(I, 13) > Max Then
    Max = ws.Cells(I, 13)
    ws.Cells(4, 17) = Max
    ws.Cells(4, 16) = ws.Cells(I, 10).Value
    End If
Next I

'Tables are being formatted to present the data in the correct conditions
ws.Range("J:M").EntireColumn.AutoFit
ws.Range("k2", ws.Range("K2").End(xlDown)).NumberFormat = "0.00"
ws.Range("L2", ws.Range("L2").End(xlDown)).NumberFormat = "0.00%"

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0.0000E+00"

'This code will employ conditional formatting to highlight negative (red) and positve (green) changes
' in stock price
For I = 2 To summaryrowcount

    If ws.Cells(I, 11).Value < 0 Then
    ws.Cells(I, 11).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(I, 11).Value > 0 Then
    ws.Cells(I, 11).Interior.ColorIndex = 4
    
    'This indicates no change in stock price for that year
    ElseIf ws.Cells(I, 11).Value = 0 Then
    ws.Cells(I, 10).Interior.ColorIndex = 7
    ws.Cells(I, 11).Interior.ColorIndex = 7
    ws.Cells(I, 12).Interior.ColorIndex = 7
    ws.Cells(I, 13).Interior.ColorIndex = 7
    
    
       
    End If
Next I

'This function will indicate which stock prices do not have an opening price
For I = 2 To summaryrowcount

    If ws.Cells(I, 14).Value = "No Open Price" Then
    ws.Cells(I, 10).Interior.ColorIndex = 1
    ws.Cells(I, 11).Interior.ColorIndex = 1
    ws.Cells(I, 12).Interior.ColorIndex = 1
    ws.Cells(I, 13).Interior.ColorIndex = 1
    ws.Cells(I, 14).Interior.ColorIndex = 1
    ws.Cells(I, 10).Font.ColorIndex = 2
    ws.Cells(I, 11).Font.ColorIndex = 2
    ws.Cells(I, 12).Font.ColorIndex = 2
    ws.Cells(I, 13).Font.ColorIndex = 2
    ws.Cells(I, 14).Font.ColorIndex = 2

    ws.Range("N:N").EntireColumn.AutoFit
    
    
    
    End If
Next I

Max = 0
Min = 0

MsgBox ("Summary for Sheet " & ws.Name & " Completed")



Next ws

End Sub




