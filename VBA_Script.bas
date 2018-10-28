Attribute VB_Name = "Module1"
' Fun begins here
Sub Analyst_Sam()

  Dim ws As Worksheet
  For Each ws In Worksheets
  ws.Activate
  Dim Ticker As String

  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Location Tracker
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Total Stock Volume"
  
  ' Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To LastRow

    ' Conditionals
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Ticker Value
      Ticker = Cells(i, 1).Value

      ' Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total_Stock_Volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

        Summary_Table_Row = Summary_Table_Row + 1
      

      Total_Stock_Volume = 0

    Else

      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If
      
  Next i

 Next ws

End Sub
