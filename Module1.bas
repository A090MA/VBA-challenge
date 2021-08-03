Attribute VB_Name = "Module1"
Sub stock_price()
 For Each ws In Worksheets
' Created a Variable to Hold File Name, Last Row, Last Column, and Year
  Dim WorksheetName As String

' Grabbed the WorksheetName
  WorksheetName = ws.Name
  ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim vol As Double
  vol = 0

  'Name the col
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "YearlyChange"
  ws.Cells(1, 11).Value = "PercentChange"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ' Name the col for Bonus
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"

  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Set last row
  Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set open_price for the first ticker
  open_price = ws.Cells(2, 3).Value
  'Dim open_price As Double
  'Dim close_price As Double
  
  ' Loop through all stock price
  For i = 2 To lastrow
    ' Add to the vol
    vol = vol + ws.Cells(i, 7).Value
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker
      ticker = ws.Cells(i, 1).Value
      ' Set the close price
      close_price = ws.Cells(i, 6).Value
      
      'Set yearly change
      y_change = close_price - open_price
      ' Set % change
      p_change = y_change / (open_price + 0.00000000001)

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker
      ' Print the anuanl change
      ws.Range("J" & Summary_Table_Row).Value = y_change
      ' Print the % change
      ws.Range("K" & Summary_Table_Row).Value = p_change * 100 & "%"
      ' Print the vol to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = vol
      
      ' Formatting
      If y_change >= 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the vol
      vol = 0
      ' Set open price
      open_price = ws.Cells(i + 1, 3).Value

    
    End If

  Next i
 
 ' Bonus

 max_inc = WorksheetFunction.Max(ws.Range("k2 : k" & Summary_Table_Row))
 min_inc = WorksheetFunction.Min(ws.Range("k2 : k" & Summary_Table_Row))
 max_vol = WorksheetFunction.Max(ws.Range("L2 : L" & Summary_Table_Row))
 ' Find the row num
 m_max_inc = WorksheetFunction.Match(max_inc, ws.Range("k2:k" & Summary_Table_Row), 0)
 m_min_inc = WorksheetFunction.Match(min_inc, ws.Range("k2:k" & Summary_Table_Row), 0)
 m_max_vol = WorksheetFunction.Match(max_vol, ws.Range("L2:L" & Summary_Table_Row), 0)
 
 ' Print
 ws.Cells(2, 17).Value = max_inc * 100 & "%"
 ws.Cells(3, 17).Value = min_inc * 100 & "%"
 ws.Cells(4, 17).Value = max_vol
 ws.Cells(2, 16).Value = ws.Cells(m_max_inc + 1, 9).Value
 ws.Cells(3, 16).Value = ws.Cells(m_min_inc + 1, 9).Value
 ws.Cells(4, 16).Value = ws.Cells(m_max_vol + 1, 9).Value
  
 Next ws
 
End Sub


