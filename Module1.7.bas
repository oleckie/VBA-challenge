Attribute VB_Name = "Module1"
'Part 1: Sort through the ticker symbols and have one of each placed in colum "I"
'Part 2: Find the price change by subtracting the first opening value from the last closing value of each stock
'Part 3: Find the percentage change for each stock
'Part 4: Find the total stock volume
'Part 5: Add conditional formatting green for positive change and red for negative change
'Part 6: Labels

Sub Stock_Analysis():

'Labels
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

'BONUS: Loop through all worksheets
    For Each StockWs In Worksheets
    

'Part 1: Sort through the ticker symbols and have one of each placed in colum "I"

    Dim output As Long
    Dim Total As Double
    Dim Start As Long
    Dim Finish As Long
    Dim Change As Double
    Dim Percent As Double
    Dim output As Long
    
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Total = 0
    output = 2
    Start = 2
    
        
    For i = 2 To lastrow
    
'Check if ticker column has populated
        If StockWs.Cells(i, 1).Value <> StockWs.Cells(i + 1, 1).Value Then
            StockWs.Cells(output, 9).Value = StockWs.Cells(i, 1).Value
'Check for yearly change per stock in J2
            Change = StockWs.Cells(i, 6) - StockWs.Cells(Start, 3)
            StockWs.Cells(output, 10).Value = Change
'Check for percentage change in K2
            Percent = StockWs.Cells(output, 10).Value / StockWs.Cells(Start, 3).Value
            StockWs.Cells(output, 11).NumberFormat = "0.00%"
'This specifies that at the end of the current ticker, it will stop and begin with the next ticker values
            Start = i + 1
            StockWs.Cells(output, 11).Value = Percent
'Check output to cell L2 for volume
            Total = Total + StockWs.Cells(i, 7).Value
            StockWs.Cells(output, 12).Value = Total
           
            output = output + 1
            Total = 0
'BONUS

        
        Else
            Total = Total + StockWs.Cells(i, 7).Value
           
        
  End If
    Next i

    Dim Chart_Row As Long
    For i = 2 To 290
    
        If (Cells(i, 10).Value > 0) Then
            StockWs.Cells(i, 10).Interior.ColorIndex = 4
        Else
            StockWs.Cells(i, 10).Interior.ColorIndex = 3
  
       
    End If
    Next i
    
    Next StockWs
    
        
End Sub
