Attribute VB_Name = "Module1"

Sub checkticker()

For Each ws In Worksheets

Dim ticker As String
Dim totalstockvol As Double
Dim openval As Currency
Dim closingval As Currency
Dim yearlychange As Currency
Dim percentchange As Double

'keep track of the location of each ticker entry in summary table
Dim SummaryRow As Double
SummaryRow = 2
'2.store opening and closing price for that year
'3.adding up the total stock volumn under each unique ticker
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

percentchange = 0
openval = ws.Cells(2, 3).Value
'Loop through from column 1 ticker at row 2
For i = 2 To lastrow


'scan if the next cell value if different from previous cell
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    totalstockvol = totalstockvol + ws.Cells(i, 7).Value
    
    closingval = ws.Cells(i, 6).Value
    
    'print the ticker symbol in summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("I" & SummaryRow).Value = ticker
    
    'print the totalstockvolume in summary table
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L" & SummaryRow).Value = totalstockvol
    
    'print yearly change in summary table
    
    yearlychange = closingval - openval
    ws.Range("J1").Value = "Yearly change"
    ws.Range("J" & SummaryRow).Value = yearlychange
        
    If yearlychange = 0 And Not IsNull(yearlychange) Then
    openval = 1
        
    Else
        percentchange = yearlychange / openval
        End If
        
    ws.Range("K1").Value = "Percentage change"
    ws.Range("K" & SummaryRow).Value = FormatPercent(percentchange)
    
    
    '---------check percent figure-------
        If yearlychange > 0 Then
            ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
        Else
            ws.Range("j" & SummaryRow).Interior.ColorIndex = 3
        End If
       
    'Add one count to summary table row
    SummaryRow = SummaryRow + 1
    
    'reset to zero
    totalstockvol = 0
    openval = ws.Cells(i + 1, 3).Value
    
    
    'Keep the value for all the beginning openingvalue
    Else
        totalstockvol = totalstockvol + ws.Cells(i, 7).Value
    End If
    Next i
    
   
   lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To lastrow2
    ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow2), 1)
    biggest = ws.Range("Q2").Value
    If ws.Cells(i, 11).Value = biggest Then
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        ws.Range("Q2") = FormatPercent(percentchange)
    
        End If
    ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow2), 1)
    smallest = ws.Range("Q3").Value
    If ws.Cells(i, 11).Value = smallest Then
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    ws.Range("Q3") = FormatPercent(percentchange)

End If
    
   ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow2), 1)
    bigtotal = ws.Range("Q4").Value
    If ws.Cells(i, 12).Value = bigtotal Then
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
    
    Next i
Next ws

End Sub





