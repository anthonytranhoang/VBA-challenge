Attribute VB_Name = "Module6"
Sub min()

    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ticker As String
    Dim i As Long
    Dim minnumber As Double

    For Each ws In Worksheets
        
        lastrow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        
        minnumber = WorksheetFunction.Max(ws.Columns(10))
        
        For i = 2 To lastrow
            If ws.Cells(i, 10).Value < minnumber Then
                minnumber = ws.Cells(i, 10).Value
                ticker = ws.Cells(i, 8).Value
            End If
        Next i
        
        ws.Cells(3, 17).Value = minnumber
        ws.Cells(3, 16).Value = ticker

    Next ws

End Sub

