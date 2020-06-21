Sub WallStreet()
' Prevents overflow error
On Error Resume Next
For Each ws In Worksheets
  ' Create headers
    ws.Cells(1, 12).Value = "ticker"
    ws.Cells(1, 13).Value = "yearly_change"
    ws.Cells(1, 14).Value = "percent_change"
    ws.Cells(1, 15).Value = "total_stock_vol"
  ' Set an intial variable for holding the last row
  Dim Last_Row As Long
  ' Determine the last row
  Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
  ' Set an initial variable for holding the ticker name
  Dim Ticker As String
  ' Set an initial variable for holding the opening price
  Dim Opening_Price As Double
  ' Set an initial variable for holding the closing price
  Dim Closing_Price As Double
  ' Set an initial variable for holding the yearly change
  Dim Yearly_Change As Double
  ' Set an initial variable for holding the percent change
  Dim Percent_Change As Double
  ' Set an initial variable for holding the total stock volume
  Dim Total_Stock_Volume As Double
  ' Set an initial variable for holding the total stock volume value
  Total_Stock_Volume = 0
  ' Keep track of the data in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  ' Loop through all data
  For i = 2 To Last_Row
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the ticker name
      Ticker = ws.Cells(i, 1).Value
      ' Sets the closing price variable
      Closing_Price = ws.Cells(i, 6).Value
      ' Sets the total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      ' Sets the yearly change variable
      Yearly_Change = Closing_Price - Opening_Price
      ' Set the percent change variable
      Percent_Change = Yearly_Change / Opening_Price * 100
      ' Print the ticker data in the summary table
      ws.Range("L" & Summary_Table_Row).Value = Ticker
      ' Print the yearly change data in the summary table
      ws.Range("M" & Summary_Table_Row).Value = Yearly_Change
      ' Print the percent change data in the summary table
      ws.Range("N" & Summary_Table_Row).Value = Percent_Change
      ' Print the total stock volume data in the summary table
      ws.Range("O" & Summary_Table_Row).Value = Total_Stock_Volume
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      ' Reset Stock Volume Total
      Total_Stock_Volume = 0
    ' Condition to determine opening price
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
      ' Sets variable for opening price
      Opening_Price = ws.Cells(i, 3)
    ' Condition to determine total stock volume
    Else
      'Sets variable for total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
    End If
Next i 
    ' Loop to apply conditional formatting
    ' Presently, this code must be run on an active worksheet to apply formatting
    For j = 2 To Summary_Table_Row
    ' Conditional for color change (green postive, red negative)
        If Cells(j, 13) >= 0 Then
        Cells(j, 13).Interior.ColorIndex = 4
        ElseIf Cells(j, 13) < 0 Then
        Cells(j, 13).Interior.ColorIndex = 3
    End If
    Next j
Next ws
End Sub
