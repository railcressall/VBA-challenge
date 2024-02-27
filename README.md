# VBA-challenge
VBA code for VBA Challenge
Sub alpha()
Dim openingprice As Double
Dim closingprice As Double
Dim percentchange As Double
Dim yearlychange As Double
Dim totalvolume As Double
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    openingprice = Cells(2, 3).Value
    closingprice = Cells(2, 6).Value
    

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
' Ticker

For i = 2 To lastRow
    Cells(i, 9).Value = Cells(i, 1)
Next i


'Yearly Change

For i = 2 To lastRow
    Cells(i, 10).Value = Cells(i, 6) - Cells(i, 3)
    
Next i

' Percent Change

 
 

End Sub
