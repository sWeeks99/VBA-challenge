Attribute VB_Name = "Module1"
Sub Stock_Data()
  
  Dim ws As Worksheet
  
  For Each ws In ThisWorkbook.Worksheets

  Dim Ticker_Name As String '(I)
  
  Dim Open_Total As Double '(JK)
  Open_Total = ws.Cells(2, "C").Value
  
  Dim Close_Total As Double '(JK)
  Close_Total = 0
  
  Dim Yearly_Change As Double '(J)
  Yearly_Change = 0
  
  Dim Percent_Change As Double '(K)
  Percent_Change = 0

  Dim Total_Volume As Double '(L)
  Total_Volume = 0
  
  Dim Summary_Row As Integer
  Summary_Row = 2
  
'For Loop start
  For i = 2 To 759001

    'Check if still within the same Ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'Set Ticker (I)
      Ticker_Name = ws.Cells(i, 1).Value

      'Set for Yearly Change (J)
      Close_Total = Close_Total + ws.Cells(i, 6).Value
      Yearly_Change = Close_Total - Open_Total
      
      ws.Cells(Summary_Row, "J").Value = Yearly_Change
      
      'Conditional Formatting: Percentage (K)
      Percent_Change = ((Close_Total / Open_Total) - 1)
      ws.Cells(Summary_Row, "K").Value = Percent_Change
      ws.Cells(Summary_Row, "K").NumberFormat = "0.00%"
      
      Open_Total = ws.Cells(i + 1, "C").Value
      Close_Total = 0
      
      'Conditional Formatting: Colors (J)
      If Yearly_Change > 0 Then
        ws.Cells(Summary_Row, "J").Interior.ColorIndex = 4
      ElseIf Yearly_Change < 0 Then
        ws.Cells(Summary_Row, "J").Interior.ColorIndex = 3
      End If

      'Add to Total Volume (L)
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      'Print the Ticker Name in Summary Row (I)
      ws.Range("I" & Summary_Row).Value = Ticker_Name

      'Print the Total Volume in Summary Row (L)
      ws.Range("L" & Summary_Row).Value = Total_Volume

      '**Add one to the summary row after finished**
      Summary_Row = Summary_Row + 1
      
      'Reset Total Volume (L)
      Total_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      'Add to Total Volume (L)
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If
     
 Next i
  
'Greatest % increase, Greatest % decrease, and Greatest total volume (OPQ)
  Dim Bottom_Row As Long
  Dim Max_Row As Long
  
  Dim Max_Percentage As Double
  Dim Min_Percentage As Double
  Dim Max_Volume As Double
  
  Dim TickerGreat As String
    
'Set Greatest % Increase and row value
    Bottom_Row = ws.Cells(Rows.Count, "K").End(xlUp).Row
    Max_Percentage = ws.Cells(2, "K").Value
    Max_Row = 2
    'For Loop for Greatest % Increase value
    For i = 2 To Bottom_Row
        If ws.Cells(i, "K").Value > Max_Percentage Then
            Max_Percentage = ws.Cells(i, "K").Value
            Max_Row = i
        End If
    Next i
    'Retrieve the Ticker Name in the row with the highest percentage rate
    TickerGreat = ws.Cells(Max_Row, "I").Value
    'Input final values
    ws.Cells(2, "P").Value = TickerGreat
    ws.Cells(2, "Q").Value = Max_Percentage
    
'Set Greatest Total Volume and row value
    Bottom_Row = ws.Cells(Rows.Count, "L").End(xlUp).Row
    Max_Volume = ws.Cells(2, "L").Value
    Max_Row = 2
    'For Loop for Greatest % Increase value
    For i = 2 To Bottom_Row
        If ws.Cells(i, "L").Value > Max_Volume Then
            Max_Volume = ws.Cells(i, "L").Value
            Max_Volume = Round(Max_Volume, 2) 'Round to the nearest hundredth place
            Max_Row = i
        End If
    Next i
    'Retrieve the Ticker Name in the row with the highest percentage rate
    TickerGreat = ws.Cells(Max_Row, "I").Value
    'Input final values
    ws.Cells(4, "P").Value = TickerGreat
    ws.Cells(4, "Q").Value = Max_Volume

'Set Greatest % Decrease and row value
    Bottom_Row = ws.Cells(Rows.Count, "K").End(xlUp).Row
    Min_Percentage = ws.Cells(2, "K").Value
    Max_Row = 2
    'For Loop for Greatest % Decrease value
    For i = 2 To Bottom_Row
        If ws.Cells(i, "K").Value < Min_Percentage Then
            Min_Percentage = ws.Cells(i, "K").Value
            Max_Row = i
        End If
    Next i
    'Retrieve the Ticker Name in the row with the lowest percentage rate
    TickerGreat = ws.Cells(Max_Row, "I").Value
    'Input final values
    ws.Cells(3, "P").Value = TickerGreat
    ws.Cells(3, "Q").Value = Min_Percentage
    
'Apply same work for other sheets
    Next ws
End Sub
