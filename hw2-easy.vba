Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Volume"
  'Set an initial variable for holding the ticker name'
  Dim ticker As String
  'Set an initial variable for holding the total volume per ticker name'
  Dim totalvolume As Double
  totalvolume = 0
  'Keep track of location for each ticker in a summary table'
  Dim summary_table_row As Integer
  summary_table_row = 2
    'Loop through all ticker volumes'
    For i = 2 To Rows.Count
      'Differentiate by ticker name'
      If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
      'Set the ticker name'
      ticker = Cells(i, 1).Value
      'Add to the totalvolume'
      totalvolume = totalvolume + Cells(i, 7).Value
      'Print ticker name to summary table'
      .Range("I" & summary_table_row).Value = ticker
      'Print volume total to summary table'
      .Range("J" & summary_table_row).Value = totalvolume
      'Add one to the summary table row'
      summary_table_row = summary_table_row + 1
      'Reset ticker total volume'
      totalvolume = 0
      'If cell immediately following row is the same ticker...'
      Else
      'Add to the total volume'
      totalvolume = totalvolume + Cells(i, 7).Value
      End If
    Next i
  Next ws
End sub
