Attribute VB_Name = "Module1"
Sub headings()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "Quarterly"
ws.Range("J1").Value = "Percent Change"
ws.Range("K1").Value = "TotalStock"
ws.Range("L1").Value = "Volume"
ws.Range("M1").Value = "OpenPrice"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"

Next ws


End Sub
