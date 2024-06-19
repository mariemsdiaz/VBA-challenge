
Sub StockAnalysis()

'Set the variables we will be working with

Dim Ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim quartelyChange As Double
Dim percentChange As Double
Dim startRow As Long
Dim nextRow As Long
Dim totalstock_volume As Double
Dim ws As Worksheet

'Loop all worksheets

For Each ws In ThisWorkbook.Worksheets

'Determine the last row for shortcut

    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Determine variables for new columns
    Dim Ticker_Column As Integer
    Ticker_Column = 2
    closePrice = 2
    total_stockvolume = 0
    startRow = 2
    
'Define new collumns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock"

'Start a loop through ticker symbols
'Calculate total stock number
'Calculate Percent Change
'Calculate Quartely Change

For i = 2 To LastRow
        total_stockvolume = total_stockvolume + ws.Cells(i, 7).Value
        Ticker = ws.Cells(i, 1).Value
        quartelyChange = ws.Cells(i, 6).Value - ws.Cells(startRow, 3)
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        closePrice = ws.Cells(i, 6).Value
        ws.Cells(2, 9).Value = Ticker
        ws.Cells(2, 10).Value = quartelyChange
        
        If ws.Cells(startRow, 3).Value <> 0 Then
        percentChange = (quartelyChange / ws.Cells(startRow, 3).Value)
        Else
        percentChange = 0
        End If
'Input Percent Change and the Total Stock Values in new columns
 
        ws.Cells(Ticker_Column, 11).Value = percentChange
   
        ws.Cells(Ticker_Column, 12).Value = total_stockvolume
    
'Move to the next row to get the rest of the results

        Ticker_Column = Ticker_Column + 1

'Reset the starting row for the next ticker
        total_stockvolume = 0
        
'Reset the starting row so it goes to the next ticker

        startRow = i + 1
       
        End If
     
Next i
Next ws


End Sub



      






