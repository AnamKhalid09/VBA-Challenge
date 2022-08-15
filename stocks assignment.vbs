Sub Stocks()
Dim openvalue As Double
Dim closevalue As Double
Dim totalvol As LongLong

'to make a seperate summary table for required results in the form of a table
Dim summary_table_row As Integer

'to run the code on all the worksheets
For Each ws In Worksheets


'declaring row number where you want to start your summary table
summary_table_row = 2

'declaring last row of each worksheet
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'to convert into percentage
ws.Columns("L").NumberFormat = "0.00%"

'to retrieve the first open value of the first stock
openvalue = ws.Cells(2, 3).Value

For i = 2 To lastrow

'to get the ticker of each stock
If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
ticker = ws.Cells(i, 1).Value

'to get the last closing value for each stock at the end of the year
closevalue = ws.Cells(i, 6).Value

'to put the ticker name in the summary table
ws.Range("J" & summary_table_row).Value = ticker

yearlychange = closevalue - openvalue
'to put the yearly change in summary table
ws.Range("K" & summary_table_row).Value = yearlychange

'to put the percentage change in summary table row
percentchange = yearlychange / openvalue
ws.Range("L" & summary_table_row).Value = percentchange

'to get the total volume of each stock and putting it in summary table
totalvol = totalvol + ws.Cells(i, 7).Value
ws.Range("M" & summary_table_row).Value = totalvol

'going to the next row of summary table
summary_table_row = summary_table_row + 1

'making totalvol equals to 0 to start the next iteration for the next row
totalvol = 0

'to get the open value for the second ticker
openvalue = ws.Cells(i + 1, 3).Value

Else
'to add up the total volume of all the rows for each ticker
totalvol = totalvol + ws.Cells(i, 7).Value
End If

'conditional formatting of cells with positive values as green and negative ones as red
If (ws.Cells(i, 11).Value < 0) Then
ws.Cells(i, 11).Interior.ColorIndex = 3
Else
ws.Cells(i, 11).Interior.ColorIndex = 4
End If

'for next iteration
Next i

'to execute same code on next worksheet
Next ws

End Sub