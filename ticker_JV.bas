Attribute VB_Name = "Module1"
Sub ticker()

  Dim ws As Worksheet
  
  ' Loop through all sheets
  For Each ws In ActiveWorkbook.Worksheets
  
  Dim lastRow As Long
  
  ' Find the last row in worksheet
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' MsgBox lastRow

  ' Set an initial variable for holding the ticker symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total stock volume per ticker symbol
  Dim Ticker_Total As Double
  Ticker_Total = 0
  
  Dim Begin_Open As Double
  Dim End_Close As Double
  
  Dim YearlyChg As Double
  Dim PercentChg As Double
  
  Dim Summary_Table_Row As Integer
  
  ' Keep track of the location for each ticker symbol in the summary table
  Summary_Table_Row = 2
  
  ws.Range("I1") = "Ticker"
  ws.Range("J1") = "Yearly Change"
  ws.Range("K1") = "Percent Change"
  ws.Range("L1") = "Total Stock Volume"
  
  ' Loop through all ticker rows
  For I = 2 To lastRow
  
    ' Check if this is a new ticker symbol
    If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
      Begin_Open = ws.Cells(I, 3)
    Else
  
    ' Check if we are still within the same ticker symbol
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the Ticker Symbol
      Ticker_Symbol = ws.Cells(I, 1).Value
      End_Close = ws.Cells(I, 6)

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + ws.Cells(I, 3).Value

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      
      YearlyChg = End_Close - Begin_Open
                  
      ' Print the Yearly Change in the Summary Table and color code
      ws.Range("J" & Summary_Table_Row).Value = YearlyChg
      
      If YearlyChg > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      If Begin_Open > 0 Then
        PercentChg = YearlyChg / Begin_Open
      Else
        PercentChg = 0
      End If
      
      ' Print the Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = PercentChg
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
      ' Print the Ticker Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
      
      ' ws.Range("N" & Summary_Table_Row).Value = Begin_Open
      ' ws.Range("O" & Summary_Table_Row).Value = End_Close
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0
      
    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + ws.Cells(I, 7).Value


    End If
    
    End If
    
  Next I
  
  ' Print the Greatest Table
  Dim Chg_Max As Double
  Dim Chg_Min As Double
  Dim TotalVol As Double
  Dim SRow As Integer
  
  ws.Range("P1") = "Ticker"
  ws.Range("Q1") = "Value"
  ws.Range("O2") = "Greatest % increase"
  ws.Range("O3") = "Greatest % decrease"
  ws.Range("O4") = "Greatest total volume"
    
  Chg_Max = Application.WorksheetFunction.Max(ws.Range("K:K"))
  Chg_Min = Application.WorksheetFunction.Min(ws.Range("K:K"))
  TotalVol = Application.WorksheetFunction.Max(ws.Range("L:L"))
  
  ws.Range("Q2").Value = Chg_Max
  ws.Range("Q2").NumberFormat = "0.00%"
  
  ws.Range("Q3").Value = Chg_Min
  ws.Range("Q3").NumberFormat = "0.00%"
  
  ws.Range("Q4").Value = TotalVol
 
  ' ws.Range("P2").Value = Application.WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K:K"), 0)
  SRow = Application.WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K:K"), 0)
  ws.Range("P2").Value = ws.Range("I" & SRow).Value
  
  ' ws.Range("P3").Value = Application.WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K:K"), 0)
  SRow = Application.WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K:K"), 0)
  ws.Range("P3").Value = ws.Range("I" & SRow).Value
  
  ' ws.Range("P4").Value = Application.WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L:L"), 0)
  SRow = Application.WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L:L"), 0)
  ws.Range("P4").Value = ws.Range("I" & SRow).Value
  
  ' Autofit to display data
  ws.Columns("I:Q").AutoFit
  
  Next ws

End Sub


