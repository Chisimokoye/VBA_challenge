Sub stockdata():
For Each ws In Worksheets

Dim i As Long
Dim Ticker As String
Dim volume As Single
volume = 0

Dim yropen As Double
Dim yrclose As Double
yropen = 0
yrclose = 0

Dim yrchange As Double
Dim percent As Double
Dim tickerrow As Long
tickerrow = 2

Dim lastrow As Single

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("K1").Value = "Percent Change"


For i = 2 To lastrow
 
 If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
 yropen = ws.Cells(i, 3).Value
 End If
 

 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    yrclose = ws.Cells(i, 6).Value
    volume = volume + ws.Cells(i, 7).Value
    yrchange = yrclose - yropen

    If yropen <> 0 Then
    percent = yrchange / yropen
    Else
    percent = 0
    End If

    
    
    If yrchange > 0 Then
    ws.Range("J" & tickerrow).Interior.ColorIndex = 4
    ElseIf yrchange <= 0 Then
    ws.Range("J" & tickerrow).Interior.ColorIndex = 3
    End If

    
    ws.Range("I" & tickerrow).Value = Ticker
    ws.Range("J" & tickerrow).Value = yrchange
    ws.Range("L" & tickerrow).Value = volume
    ws.Range("K" & tickerrow).Value = percent
    ws.Range("K" & tickerrow).NumberFormat = "0.00%"

    tickerrow = tickerrow + 1
    volume = 0
    Else
    volume = volume + Cells(i, 7).Value
    
    
    
    
    End If
    
    Next i
    
'Bonus

ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))


ws.Range("P1").Value = "Ticker Symbol"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

Dim increase As Double
Dim decrease As Double
Dim t_volume As Single
increase = ws.Range("Q2").Value
decrease = ws.Range("Q3").Value
t_volume = ws.Range("Q4").Value

For j = 2 To lastrow
If ws.Cells(j, 11).Value = increase Then
ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
End If

If ws.Cells(j, 11).Value = decrease Then
ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
End If

If ws.Cells(j, 12).Value = t_volume Then
ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
End If

Next j


Next ws


End Sub

