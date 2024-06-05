Attribute VB_Name = "Module2"
Sub greenred()
    Dim lastrow As Long
    Dim i As Long
    Dim ws As Worksheet

    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
            If ws.Cells(i, 9).Value > 0 Then
                ws.Cells(i, 9).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 9).Value < 0 Then
                ws.Cells(i, 9).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 9).Interior.ColorIndex = xlNone
            End If
        Next i
    Next ws
End Sub
