Attribute VB_Name = "Module1"
Sub openprice()
    Dim ws As Worksheet
    Dim openprice As Double
    Dim ticker As String
    Dim lastrow As Long
    Dim outputRow As Long

    For Each ws In ThisWorkbook.Worksheets
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputRow = 2 ' Start output from row 2

        For i = 2 To lastrow
            ' Check if the next row has a different ticker (or if it's the last row)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openprice = ws.Cells(i, 3).Value

                ' Output the ticker and open price in columns H and I
                ws.Cells(outputRow, 8).Value = ticker
                ws.Cells(outputRow, 12).Value = openprice
                
                outputRow = outputRow + 1 ' Move to the next output row
            End If
        Next i
    Next ws
End Sub
     
    
