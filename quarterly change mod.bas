Attribute VB_Name = "Module1"
Sub quartchange()

  For Each ws In Worksheets
  
  Dim ticker As String

  Dim quartchange As Double
  quartchange = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
  For i = 2 To lastrow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      ticker = Cells(i, 1).Value

      ' Add to the Brand Total
      quartchange = Cells(i, 3).Value - Cells(i, 6).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("H" & Summary_Table_Row).Value = ticker

      ' Print the Brand Amount to the Summary Table
      Range("I" & Summary_Table_Row).Value = quartchange

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      quartchange = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      quartchange = quartchange + Cells(i, 3).Value

    End If

  Next i
  
  Next ws
  
End Sub

