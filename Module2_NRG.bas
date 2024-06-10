Attribute VB_Name = "Module1"
Sub Module2()

' git link: https://github.com/natbiorg/VBA_challenge.git

'Learned Application.ScreenUpdating from ChatGPT to ensure code runs smoothly/faster
Application.ScreenUpdating = False

'I used ChatGPT to help debug as I went through the code. Asking questions like: "I ran my code but received an overflow error" or "How do I format Total Stock volume to not include E? "
'ChatGPT also helped me determine that row should be Long instead of Integer and which variables to be Double
'I worked on this Module in person with Lily O'Connel

'make this applicable to any worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets

    
    ' defining variables and setting them to their base value
    Dim row As Long
    row = 2
    Dim lastrow As Long
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Dim ssrow As Long
    ssrow = 2
    
    Dim conditionalrow As Long
    conditionalrow = ws.Cells(ws.Rows.Count, "J").End(xlUp).row
    
    Dim POpenSum As Double
    Dim PCloseSum As Double
    Dim QChange As Double
    Dim PercentQChange As Double
    Dim TotalStock As Double
    POpenSum = 0
    PCloseSum = 0
    
    Dim PQChangeInc As Double
    Dim PQChangeDec As Double
    Dim TotalStockMax As Double
    Dim TickerMax As String
    Dim TickerMin As String
    Dim TickerTotVol As String
    PQChangeInc = 0
    PQChangeDec = 0
    TotalStockMax = 0
    
    Dim rng As Range
    Set rng = ws.Range("J2:J" & lastrow)
    
    'Placing Headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Quarterly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'first loop that goes row by row and pulls values into created columns
    For row = 2 To lastrow
    
        'Creating a loop if ticker = ticker + 1
        If ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value Then
            POpenSum = POpenSum + ws.Cells(row, 3).Value
            PCloseSum = PCloseSum + ws.Cells(row, 6).Value
            TotalStock = TotalStock + ws.Cells(row, 7).Value
        
        Else
            'assigning ticker to summary row
            ws.Cells(ssrow, 9).Value = ws.Cells(row, 1).Value
            
            'calculating and assigning quarterly change - rounding so that it looks reasonable
            QChange = POpenSum - PCloseSum
            QChange = Round(QChange, 3)
            ws.Cells(ssrow, 10).Value = QChange
        
            'calculating and assigning percentage change
            PercentQChange = (POpenSum - PCloseSum) / POpenSum
            ws.Cells(ssrow, 11).Value = PercentQChange
            ws.Cells(ssrow, 11).NumberFormat = "0.00%"
            
            'calculating and assigning total stock volume
            ws.Cells(ssrow, 12).Value = TotalStock
            ws.Cells(ssrow, 12).NumberFormat = "#,##0"
            
            'reseting for next ticker
            ssrow = ssrow + 1
            POpenSum = 0
            PCloseSum = 0
            TotalStock = 0
        
        End If
        
        'calculating and outputting the stock with the greatest % increase, etc
        'Greatest Increase
        If PercentQChange > PQChangeInc Then
            PQChangeInc = PercentQChange
            TickerMax = ws.Cells(row, 1).Value
        Else
            PQChangeInc = PQChangeInc
            ws.Cells(2, 17).Value = PQChangeInc
            ws.Cells(2, 16).Value = TickerMax
            ws.Cells(2, 17).NumberFormat = "0.00%"
        End If
        
        'Greatest Decrease
        If PercentQChange < PQChangeDec Then
            PQChangeDec = PercentQChange
            TickerMin = ws.Cells(row, 1).Value
        Else
            PQChangeDec = PQChangeDec
            ws.Cells(3, 17).Value = PQChangeDec
            ws.Cells(3, 16).Value = TickerMin
            ws.Cells(3, 17).NumberFormat = "0.00%"
        End If
        
        'Greatest Stock
        If TotalStock > TotalStockMax Then
            TotalStockMax = TotalStock
            TickerTotVol = ws.Cells(row, 1).Value
        Else
            TotalStockMax = TotalStockMax
            ws.Cells(4, 17).Value = TotalStockMax
            ws.Cells(4, 16).Value = TickerTotVol
            ws.Cells(4, 17).NumberFormat = "#,##0"
        End If
        
    Next row

'**** conditional formatting
'**** this conditional formatting was developed using Xpert Learning Assistant, please see **** at end of code

    rng.FormatConditions.Delete
    ' Apply conditional formatting for negative values (Red color)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    .Interior.Color = RGB(255, 0, 0) ' Red color
    End With

' Apply conditional formatting for positive values (Green color)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    .Interior.Color = RGB(0, 255, 0) ' Green color
    End With


Next ws

Application.ScreenUpdating = True

'**** I initially approached conditional formatting by trying to use it in the row loop, but this kept formatting column J in the first few sheets as green and the last few sheets as red without following conditions
 'If QChange > 0 Then
        'ws.Cells(row, 10).Interior.ColorIndex = 4
 'ElseIf QChange < 0 Then
        'ws.Cells(row, 10).Interior.ColorIndex = 3
 'Else
        'ws.Cells(row, 10).Interior.ColorIndex = xlNone
' End If

'****So i tried asking Gemini which returned the conditional formatting code below, again this code would not format according to conditions - it would return results only for the first sheet and the formatting was incorrect**
    'With ws.Range("J2:J" & lastrow)  ' Assuming QChange is in column J
  '.FormatConditions.Add Type:=xlExpression, Formula1:="=J2>0"
    '.Interior.ColorIndex = 4  ' Green for positive
  '.FormatConditions.Add Type:=xlExpression, Formula1:="=J2<0"
    '.Interior.ColorIndex = 3  ' Red for negative
   ' End With

End Sub
