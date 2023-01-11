Sub Stock_Data()

Dim ws As Worksheet

' Loop through all Sheets'
    For Each ws In Worksheets
    
' Assign headers for the Ticker table'
 
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percentage Change"
 ws.Range("L1").Value = "Total Stock Volume"
 
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
 ws.Range("O2").Value = "Greatest % Increase"
 ws.Range("O3").Value = "Greatest % Decrease "
 ws.Range("O4").Value = "Greatest Total Volume"
 
'-----------------------------'
 Dim Ticker_name As String
 Dim OpenPrice As Double
 OpenPrice = 0
 Dim ClosePrice As Double
 ClosePrice = 0
 Dim YearlyChange As Double
 YearlyChange = 0
 Dim PercentChange As Double
 PercentChange = 0
 Dim TotalVolume As LongLong
 Dim Firstrow As Long
 Firstrow = 2
 
 ' Keep track of each Ticker for Ticker table row'
 
 Dim Ticker_Table_Row As Integer
 
 Ticker_Table_Row = 2
 
 
 ' Determine the last row '
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
' Loop through Ticker'

For i = 2 To LastRow

' Check if Ticker name is same, if not then'

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
Ticker_name = ws.Cells(i, 1).Value

' Print the Ticker name in the Ticker table'

ws.Range("I" & Ticker_Table_Row).Value = Ticker_name

' Add 1 to Ticker table row'
Ticker_Table_Row = Ticker_Table_Row + 1
Firstrow = Firstrow + 1
YearlyChange = 0

' Calculate Yearly Change'
ClosePrice = ws.Cells(i, 6).Value
OpenPrice = ws.Cells(Firstrow, 3).Value
YearlyChange = ClosePrice - OpenPrice
 
ws.Range("J" & Ticker_Table_Row).Value = YearlyChange
 

 ' Calculate Percentage Change'
  ElseIf OpenPrice <> 0 Then
  
  PercentChange = (YearlyChange / OpenPrice) * 100
  ws.Range("K" & Ticker_Table_Row).Value = PercentChange
  
  End If
  
  ' Calculate Total volume'
  
  TotalVolume = 0
  
  TotalVolume = TotalVolume + ws.Cells(i, 7).Value
  ws.Range("L" & Ticker_Table_Row).Value = TotalVolume
   
 ' Conditional coloring of Percentage Change'
 
If ws.Cells(i, 11) >= 0 Then

' Color green when positive'

  ws.Cells(i, 11).Interior.ColorIndex = 4
  
  Else
  
    ws.Cells(i, 11).Interior.ColorIndex = 3
  

End If


 Next i
 
 'Find Greatest % Increase,Greatest % Decrease & Greatest Total Volume'
 
 Dim Max_Increase As Double
 Dim Max_Decrease As Double
 Dim Max_TotalVolume As LongLong
 Dim Ticker_Index_MaxIncrease As String
 Dim Ticker_Index_MaxDecrease As String
 Dim Ticker_Index_MaxTotalVolume As String

 Max_Increase = WorksheetFunction.Max(Range("K2:K" & LastRow))
 ws.Range("Q2").Value = Max_Increase

 Max_Decrease = WorksheetFunction.Min(Range("K2:K" & LastRow))
 ws.Range("Q3").Value = Max_Decrease
 
 Max_TotalVolume = WorksheetFunction.Max(Range("L2:L" & LastRow))
 ws.Range("Q4").Value = Max_TotalVolume
 
 Ticker_Index_MaxIncrease = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Max_Increase, Range("K2:K" & LastRow), 0))
 Ticker_Index_MaxDecrease = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Max_Decrease, Range("K2:K" & LastRow), 0))
 Ticker_Index_MaxTotalVolume = WorksheetFunction.Index(Range("I2:I" & LastRow), WorksheetFunction.Match(Max_TotalVolume, Range("L2:L" & LastRow), 0))
 
 ws.Range("P2").Value = Ticker_Index_MaxIncrease
 ws.Range("P3").Value = Ticker_Index_MaxDecrease
 ws.Range("P4").Value = Ticker_Index_MaxTotalVolume
 

 Next ws


End Sub
