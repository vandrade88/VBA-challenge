Sub stock_market()

' Declare and set worksheet
Dim ws As Worksheet

' Loop through all stocks for one year
For Each ws In Worksheets

' Create the column headings
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 12).Value = "Total Stock Volume"

' Create the column headings
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"

' Set an initial variable for holding the first day open amount
Dim first_day_open As Double

' Calculate initial first day open value of stock
first_day_open = ws.Cells(2, 3).Value

' Set an initial variable for holding the yearly change amount
Dim yearly_change As Double

' Set an initial variable for holding the last day close amount
Dim last_day_close As Double

' Set an initial variable for holding the percent change of the yearly change
Dim percent_change As Double

' Set an initial variable for holding the ticker name
Dim ticker_name As String

' Set an initial variable for holding the total stock volume amount
Dim total_stock_volume As Double
total_stock_volume = 0

' Keep track of the location for each ticker name in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

' Determine the last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Make a for loop of current worksheet to last row
For i = 2 To lastrow

' If else conditional to check ticker symbol names and stock volume
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Print the ticker name
ticker_name = ws.Cells(i, 1).Value

' Add to the total stock volume amount
total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

' Print the ticker name into the summary table
ws.Range("I" & summary_table_row).Value = ticker_name

' Print the total stock volume amount into the summary table
ws.Range("L" & summary_table_row).Value = total_stock_volume

' Calculations for yearly change and percent change
last_day_close = ws.Cells(i, 6).Value
yearly_change = last_day_close - first_day_open

' Print the yearly price change in the summary table
ws.Range("J" & summary_table_row).Value = yearly_change

' Conditional formatting to highlight positive change in green and negative change in red
If yearly_change > 0 Then
ws.Range("J" & summary_table_row).Interior.ColorIndex = 4

ElseIf yearly_change <= 0 Then
ws.Range("J" & summary_table_row).Interior.ColorIndex = 3

End If

' Condition for zero value error
    If first_day_open = 0 Then
    percent_change = 0

    Else
    percent_change = yearly_change / first_day_open
    
    End If

' Print the percent change amount as a percentage in the summary table
ws.Range("K" & summary_table_row).Value = percent_change
ws.Range("K" & summary_table_row).NumberFormat = "0.00%"

' Add one to the summary table row
summary_table_row = summary_table_row + 1
      
' Reset the total stock volume amount
total_stock_volume = 0

' Reset the next stock's open price
first_day_open = ws.Cells(i + 1, 3).Value

Else

' Add to the total stock volume amount
total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub