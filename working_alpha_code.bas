Attribute VB_Name = "Module1"
Sub alpha()
Dim ws As Worksheet
Dim lastrow As Long
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim total As Double
Set ws = ThisWorkbook.Sheets("A")
For Each ws In ActiveWorkbook.Sheets

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> "" Then
        ticker = ws.Cells(i, 1).Value
        opening = ws.Cells(i, 3).Value
        closing = ws.Cells(i, 6).Value
        yearlychange = closing - opening
        percentchange = ((closing - opening) / opening) * 100
        total = ws.Cells(i, 7).Value
        
        ws.Cells(i, 9) = ticker
        ws.Cells(i, 10) = yearlychange
        ws.Cells(i, 11) = percentchange
        ws.Cells(i, 12) = total
    
        End If

    Next i
    
For i = 2 To lastrow

    If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
    
    End If

    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If

Next i
Next ws


End Sub
