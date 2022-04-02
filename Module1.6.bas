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
    
'Part 1: Sort through the ticker symbols and have one of each placed in colum "I"

    Dim output As Long
    Dim Total As Double
    Dim Start As Long
    Dim Finish As Long
    Dim Change As Double
    Dim Percent As Double
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Total = 0
    output = 2
    Start = 2
    
        
    For i = 2 To lastrow
    
'Check if ticker column has populated
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(output, 9).Value = Cells(i, 1).Value
'Check for yearly change per stock in J2
            Change = Cells(i, 6) - Cells(Start, 3)
            Cells(output, 10).Value = Change
'Check for percentage change in K2
            Percent = Cells(output, 10).Value / Cells(Start, 3).Value
            Cells(output, 11).NumberFormat = "0.00%"
'This specifies that at the end of the current ticker, it will stop and begin with the next ticker values
            Start = i + 1
            Cells(output, 11).Value = Percent
'Check output to cell L2 for volume
            Total = Total + Cells(i, 7).Value
            Cells(output, 12).Value = Total
           
            output = output + 1
            Total = 0
        Else
            Total = Total + Cells(i, 7).Value
           
        
  End If
    Next i

        
    For i = 2 To 290
    
        If (Cells(i, 10).Value > 0) Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
  
       
    End If
    Next i
    
         
        
End Sub
