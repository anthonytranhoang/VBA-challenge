Attribute VB_Name = "Module4"
Sub percent()
    Dim ws As Worksheet
    Dim percentchange As Double
    Dim openprice As Double
    Dim lastrow As Long
    Dim outputrow As Long

    For Each ws In ThisWorkbook.Worksheets
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputrow = 2
        
        For i = 2 To lastrow
            If ws.Cells(i, 12).Value <> 0 Then
                percentchange = ((ws.Cells(i, 9).Value / ws.Cells(i, 12).Value) * 100)
            Else
                percentchange = 0
            End If
            
            ws.Cells(i, 10).Value = percentchange
        Next i
    Next ws
End Sub
