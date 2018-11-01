Sub ticker_vol()

' Set an initial variable for holding the Ticker name
  Dim Ticker_Name As String

   ' Determine the Last Row
   For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
  ' Set an initial variable for holding the total volume per ticker
  Dim ticker_Total As Double
  ticker_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Variables for Yearly change and percent change
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
  
   ws.Range("I1").Value = "Ticker"
   ws.Range("I1").Font.ColorIndex = 1
   ws.Range("I1").Font.Bold = True
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("J1").Font.ColorIndex = 1
   ws.Range("J1").Font.Bold = True
   ws.Range("K1").Value = "Percent Change"
   ws.Range("K1").Font.ColorIndex = 1
   ws.Range("K1").Font.Bold = True
   ws.Range("L1").Value = "Total Stock Volume"
   ws.Range("L1").Font.ColorIndex = 1
   ws.Range("L1").Font.Bold = True
   
  ' Set Open Price
  Open_Price = ws.Cells(2, 3).Value
  ' Loop through all tickers
  For i = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker Total
      ticker_Total = ticker_Total + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Volume Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = ticker_Total

    ' Set Close Price
                Close_Price = ws.Cells(i, 6).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("J" & Summary_Table_Row).NumberFormat = "0.000000000000000"
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                End If
                
                If Yearly_Change >= 0 Then
                  ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Reset OPen Price
      Open_Price = ws.Cells(i + 1, 3).Value
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



