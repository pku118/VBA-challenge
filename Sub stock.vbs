Sub stock()

  ' Set an initial variable for holding the brand name
  Dim Stock_Tic As String
  
  'counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Set an initial variable for holding the total Stock Volume
  Dim Stock_VolTot As Double
  Dim Stock_Open As Double
  Dim Stock_Close As Double
  
  Stock_VolTot = 0
  Stock_Open = 0
  Stock_Close = 0

  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Stocks
  For i = 2 To lastrow

    ' Check if we are still within the same Stock Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Stock Tic
      Stock_Tic = Cells(i, 1).Value
      
      'Set the Stock Open
      Stock_Open = Cells(i, 3).Value

      ' Add to the Volume Total
      Stock_VolTot = Stock_VolTot + Cells(i, 7).Value

      ' Print the Stock Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Stock_Tic

      'Print the Open Price
      Range("M" & Summary_Table_Row).Value = Stock_Open

      ' Print the Stock Vol Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Stock_VolTot

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Stock_VolTot = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Stock_VolTot
      Stock_VolTot = Stock_VolTot + Cells(i, 7).Value
      Stock_Close = Cells(i, 6).Value
      Range("N" & Summary_Table_Row).Value = Stock_Close
      'Stock_Close = Cells(i, 14).Value

    End If

  Next i

End Sub
