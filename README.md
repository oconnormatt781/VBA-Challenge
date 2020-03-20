# VBA-Challenge
Week 2 VBA Challenge

Sub Stocks()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Ticker
    Next
    Application.ScreenUpdating = True
End Sub
Sub Ticker():

    'Set variables for ticker, yearly change, percent change, and total volume.
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume As LongLong
    
    ' Keep track of each ticker in summary columns
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Loop through all tickers
    For i = 2 To 70927
    
' Check if still on the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set ticker
      Ticker = Cells(i, 1).Value
      
      ' Calculate yearly change
      open_price = Cells(i, 3).Value
      close_price = Cells(i, 6).Value
      Yearly_Change = (close_price - open_price)
      
      ' Calculate percent change
      Percent_Change = ((close_price - open_price) / (open_price))

      ' Add up the volume
      Volume = Volume + Cells(i, 7).Value

      ' Print the ticker
      Range("H" & Summary_Table_Row).Value = Ticker
      
      ' Print the yearly change to the summary
      Range("I" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the percent change to summary
      Range("J" & Summary_Table_Row).Value = Percent_Change

      ' Print the total volume to summary
      Range("K" & Summary_Table_Row).Value = Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker
      Ticker = 0
      'Reset Volume
      Volume = 0

    ' If the cell immediately following a row is the same ticker
    Else

      ' Add to the volume total
      Volume = Volume + Cells(i, 7).Value

    End If

  Next i

End Sub

