Sub vba_challenge()

' Create Headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"


' Set an initial variable for holding the ticker name
  Dim ticker_name As String


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

  Next i

End Sub
