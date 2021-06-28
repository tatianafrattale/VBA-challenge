Sub vba_challenge()

' Create Headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"


' Set an initial variables
  Dim ticker_name As String
  Dim last_total As Long
    last_total = 2
  Dim year_open As Double
  Dim year_close As Double
  Dim yearly_change As Double
  Dim total_stock As Long
  Dim percent_change As Double



' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' Find the last row
  lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Create Loop
  For i = 2 To lastrow

    ' Check if we are still within the same ticker name, if not...
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      ' Set the ticker name
     ticker_name = Cells(i, 1).Value


      ' Print the ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker_name


      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

    Else
    End If


      ' Create variables to hold values for yearly change and place in correct column
      year_open = Range("C" & last_total)
      year_close = Range("F" & i)
      yearly_change = year_close - year_open
      Range("J" & Summary_Table_Row).Value = yearly_change

      ' Determine percent change and place in correct column
      If year_open = 0 Then
        percent_change = 0
      Else
      percent_change = yearly_change / year_open * 100
      Range("K" & Summary_Table_Row).Value = percent_change

    End If

      ' Use conditional formatting to insert green and red on yearly change values
      If Cells(i, 10).Value <= 0 Then
         Cells(i, 10).Interior.ColorIndex = 3
      Else
         Cells(i, 10).Interior.ColorIndex = 4
      End If

    ' Create variable that holds total stock value
    total_stock = 0
    total_stock = total_stock + Cells(i, 7).Value
    Range("L" & Summary_Table_Row).Value = total_stock
    
     ' Create if to reset the stock count if the ticker is different
   
      If Cells(i + 1, 1) <> Cells(i, 1) Then
            Cells(i, 9).Value = ticker
            Cells(i, 10).Value = total_stock
            total_stock = 0
   End If

   
    
  Next i

End Sub

