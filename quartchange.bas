Attribute VB_Name = "Module2"
Sub quartchange()
    Dim ws As Worksheet
    Dim quartchange As Double
    Dim openprice As Double
    Dim lastrow As Long
    Dim outputrow As Long

    For Each ws In ThisWorkbook.Worksheets
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputrow = 2
        
        For i = 2 To lastrow
            If ws.Cells(i, 12).Value <> 0 Then
                quartchange = ws.Cells(i, 9).Value - ws.Cells(i, 12).Value
            Else
                quartchange = 0
            End If
            
            ws.Cells(i, 9).Value = quartchange
        Next i
    Next ws
End Sub


