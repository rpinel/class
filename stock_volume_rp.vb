Sub ticker_vol()

' Set an initial variable for holding the Ticker name
  Dim ticker_Name As String

   ' Determine the Last Row
   For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
  ' Set an initial variable for holding the total volume per ticker
  Dim ticker_Total As Double
  ticker_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
   ws.Range("I1").Value = "Ticker"
   ws.Range("I1").Font.ColorIndex = 1
   ws.Range("I1").Font.Bold = True
   ws.Range("J1").Value = "Total Stock Volume"
   ws.Range("J1").Font.ColorIndex = 1
   ws.Range("J1").Font.Bold = True
   
  ' Loop through all tickers
  For i = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker Total
      ticker_Total = ticker_Total + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker_Name

      ' Print the Ticker Volume Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      ticker_Total = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker Total
      ticker_Total = ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i
Next ws

End Sub

