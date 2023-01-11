Attribute VB_Name = "Module1"
Sub TickerSummary()

'Set variable types for ticker summary
Dim i As Long
Dim lastrow As Long
Dim ticker As String
Dim tickercount As Long
Dim openprice As Single
Dim closeprice As Single
Dim stockvolume As Double
Dim pricechange As Currency
Dim ws As Worksheet

'Set variable type for Greatest Summary
Dim greatestpercentinc As Single
Dim greatestperinctick As String
Dim greatestpercentdec As Single
Dim greatestperdectick As String
Dim greatestvol As Double
Dim greatestvoltick As String

'Activate each worksheet in workbook and run script on them all
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Reset stockvolume, openprice, closeprice and tickercount between each worksheet
stockvolume = 0
openprice = 0
closeprice = 0
tickercount = 0

'Set greatest variables to zero for worksheet
greatestpercentinc = 0
greatestpercentdec = 0
greatestvol = 0


'Set lastrow variable to last non-blank row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Print Headers for Columns I:J
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Run loop from row 2 to last non-blank row (lastrow)

For i = 2 To lastrow
    'Add volume values (from Column G: <vol>) together for the current ticker
    stockvolume = stockvolume + Cells(i, 7).Value
    
    'Look for the first row of the current ticker by comparing the previous row's value in Column A
    'to the current row's value. If they are different, then store the value from this row's
    'Column C (Opening stock value: <open>) in openprice variable
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    openprice = Cells(i, 3).Value
    
    'Look for the last entry for the current ticker by comparing the Column A value of the next row to the
    'Column A value of the current row. If they are different, store the ticker string in ticker variable, store
    'the close price from current row, Column F (<close>) in closeprice variable and progress tickercount variable
    'by 1 in order to ensure values are printed in the first open row of Columns I:J
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        closeprice = Cells(i, 6).Value
        tickercount = tickercount + 1
        
        ' Print ticker in first open row of Column I (9)
        Cells(1 + tickercount, 9).Value = ticker
        
        ' Print Yearly Change in first open row of Column J (10) with two decimal places and a leading zero where necessary
        Cells(1 + tickercount, 10).Value = closeprice - openprice
        Cells(1 + tickercount, 10).NumberFormat = "#0.00"
        
            ' Format Cell based on Yearly Change value. Red if negative, Green if positive or zero
            If Cells(1 + tickercount, 10).Value >= 0 Then
            Cells(1 + tickercount, 10).Interior.ColorIndex = 4
            Else: Cells(1 + tickercount, 10).Interior.ColorIndex = 3
            End If
            
        ' Print Percent Change in first open row of Column K (11)
        Cells(1 + tickercount, 11).Value = FormatPercent((closeprice - openprice) / openprice, 2)
        
        
            ' Format Cell based on Percent Change value. Red if negative, Green if positive or zero
            If Cells(1 + tickercount, 11).Value >= 0 Then
            Cells(1 + tickercount, 11).Interior.ColorIndex = 4
            Else: Cells(1 + tickercount, 11).Interior.ColorIndex = 3
            End If
        
        
            'If Percent Change is greater than stored greatestpercentinc, replace with this value
            If ((closeprice - openprice) / openprice) > greatestpercentinc Then
            greatestpercentinc = (closeprice - openprice) / openprice
            greatestperinctick = ticker
            End If
        
            'If Percent Change is lower than stored greatestpercentdec, replace with this value
            If ((closeprice - openprice) / openprice) < greatestpercentdec Then
            greatestpercentdec = (closeprice - openprice) / openprice
            greatestperdectick = ticker
            End If
        
        
        ' Print Total Annual Stock Volume in first open row of Column L (13)
        Cells(1 + tickercount, 12).Value = stockvolume
        
            'If Total Stock Volumme is greater than stored greatestvol, replace with this value
            If stockvolume > greatestvol Then
            greatestvol = stockvolume
            greatestvoltick = ticker
            End If
        
        'Reset stockvolume, openprice and closeprice to zero
        stockvolume = 0
        openprice = 0
        closeprice = 0
  
  End If
  Next i
  
'Autofit Columns I:L to Header String Length
Columns("I:L").EntireColumn.AutoFit


'Add headers for Columns P:Q
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Add Labels for Greatest % Increase, % Decrease and Total Volume
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Print Greatest Tickers with respective values for each of the above categories
Range("P2").Value = greatestperinctick
Range("Q2").Value = FormatPercent(greatestpercentinc, 2)
Range("P3").Value = greatestperdectick
Range("Q3").Value = FormatPercent(greatestpercentdec, 2)
Range("P4").Value = greatestvoltick
Range("Q4").Value = greatestvol

'Autofit Columns O:Q to Header String Length
Columns("O:Q").EntireColumn.AutoFit

    
'Progress to next worksheet in current workbook
Next ws


End Sub


