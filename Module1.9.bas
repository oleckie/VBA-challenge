Attribute VB_Name = "Module1"
'Part 1: Sort through the ticker symbols and have one of each placed in colum "I"
'Part 2: Find the price change by subtracting the first opening value from the last closing value of each stock
'Part 3: Find the percentage change for each stock
'Part 4: Find the total stock volume
'Part 5: Add conditional formatting green for positive change and red for negative change
'Part 6: Labels

Sub Stock_Analysis():

'BONUS: LOOP
For Each ws In Worksheets


'Labels
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
'Part 1: Sort through the ticker symbols and have one of each placed in colum "I"

    Dim output As Long
    Dim Total As Double
    Dim Start As Long
    Dim Finish As Long
    Dim Change As Double
    Dim Percent As Double
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Total = 0
    output = 2
    Start = 2
    
        
    For i = 2 To lastrow
    
'Check if ticker column has populated
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(output, 9).Value = ws.Cells(i, 1).Value
'Check for yearly change per stock in J2
            Change = ws.Cells(i, 6) - ws.Cells(Start, 3)
            ws.Cells(output, 10).Value = Change
'Check for percentage change in K2
        If ws.Cells(Start, 3).Value = 0 Then
            Percent = 0
        Else
            Percent = ws.Cells(output, 10).Value / ws.Cells(Start, 3).Value
            ws.Cells(output, 11).NumberFormat = "0.00%"
        End If
        
'This specifies that at the end of the current ticker, it will stop and begin with the next ticker values
            Start = i + 1
            ws.Cells(output, 11).Value = Percent
'Check output to cell L2 for volume
            Total = Total + ws.Cells(i, 7).Value
            ws.Cells(output, 12).Value = Total
           
            output = output + 1
            Total = 0
        Else
            Total = Total + ws.Cells(i, 7).Value
 
        
  End If
    Next i

'Conditional Formatting
    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
    For i = 2 To lastrow
    
        If (ws.Cells(i, 10).Value > 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
  
       
    End If
    Next i
    
'BONUS: Greatest % Increase & Decrease
'Labels
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest & Decrease"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
  
    Dim percentend As Long
    Dim percentmax As Double
    Dim percentmin As Double
    
    percentend = ws.Cells(Rows.Count, 11).End(xlUp).Row
    percentmax = 0
    percentmin = 0
    
    For i = 2 To percentend
    
    If percentmax < ws.Cells(i, 11).Value Then
        percentmax = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = percentmax
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
    ElseIf percentmin > ws.Cells(i, 11).Value Then
        percentmin = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = percentmin
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
    End If
    Next i

'BONUS: Greatest Total Volume
ws.Cells(4, 15) = "Greatest Total Volume"
Dim volumeend As Long
Dim volumemax As Double

volumeend = ws.Cells(Rows.Count, 12).End(xlUp).Row
volumemax = 0

For i = 2 To volumeend

    If volumemax < ws.Cells(i, 12).Value Then
        volumemax = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = volumemax
        
    End If
Next i

 Next ws
    
                 
End Sub

