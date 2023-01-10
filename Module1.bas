Attribute VB_Name = "Module1"
Sub solution()

For Each ws In Worksheets

Dim ticker As String
Dim yearly_change As Double
Dim percent_changes As Double
Dim total_stock As Double
Dim Summary_Table_Row As Integer
Dim LastRow As Long

Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

total_stock = 0

opening = ws.Cells(2, 3).Value

  ' Loop through all tickers
  For i = 2 To LastRow

    ' Check if we are still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      ticker = ws.Cells(i, 1).Value
      
      'calculate yearly change and percentage change
      'opening = ws.Cells(i, 3).Value
      closing = ws.Cells(i, 6).Value
      yearly_change = closing - opening
      percent_change = FormatPercent(((closing / opening) - 1))
      
      opening = ws.Cells(i + 1, 3).Value

      ' Add to the stock volume total
      total_stock = total_stock + ws.Cells(i, 7).Value
      
      'set the headers
      ws.Range("I1").Value = "ticker name"
      ws.Range("J1").Value = "yearly change"
      ws.Range("K1").Value = "percentage change"
      ws.Range("L1").Value = "total stock volume"

      ' Print the information in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      ws.Range("K" & Summary_Table_Row).Value = percent_change
      ws.Range("L" & Summary_Table_Row).Value = total_stock

      'add conditional formatting
    If yearly_change > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock
      total_stock = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the total stock volume
      total_stock = total_stock + ws.Cells(i, 7).Value
      
      End If
      
    Next i
    
    'find max percent increase and decrease
    max_percentage_increase = WorksheetFunction.Max(ws.Range("K:K"))
    max_percentage_decrease = WorksheetFunction.Min(ws.Range("K:K"))
    greatest_stock_val = WorksheetFunction.Max(ws.Range("L:L"))
    
    ws.Range("N2").Value = "max percentage increase"
    ws.Range("N3").Value = "max percentage decrease"
    ws.Range("O2").Value = FormatPercent(max_percentage_increase)
    ws.Range("O3").Value = FormatPercent(max_percentage_decrease)
    ws.Range("N4").Value = "greatest total stock"
    ws.Range("o4").Value = greatest_stock_val
    
    
Next ws

End Sub



