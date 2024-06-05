Attribute VB_Name = "Module2"
Sub volume()

    Dim ws As Worksheet
    Dim ticker As String
    Dim totalvol As Double
    totalvol = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim lastrow As Long
    Dim i As Long


    For Each ws In ThisWorkbook.Worksheets
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
     
        For i = 2 To lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                
                
                totalvol = totalvol + ws.Cells(i, 7).Value
                
               
                ws.Range("H" & Summary_Table_Row).Value = ticker
                
               
                ws.Range("K" & Summary_Table_Row).Value = totalvol
                
                
                Summary_Table_Row = Summary_Table_Row + 1
                
               
                totalvol = 0
            Else
                
                totalvol = totalvol + ws.Cells(i, 7).Value
            End If
        Next i
        
      
        Summary_Table_Row = 2
    Next ws

End Sub
