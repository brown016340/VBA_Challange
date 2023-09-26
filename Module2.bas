Attribute VB_Name = "Module2"
Sub SummarizeAllYears()

'runs code on all worksheets
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub


Sub RunCode()
    
'Headers
Cells(1, 8).Value = "Ticker"
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Total Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Cells(2, 14).Value = "Greatest % Increse"
Cells(3, 14).Value = "Greatest % Decrese"
Cells(4, 14).Value = "Greatest Total Volume"


'defining variables
Dim ticker As String
Dim totalVolume As Double
Dim yearlyChange As Double
Dim summaryRow As Integer
  summaryRow = 1
Dim yearStart As Double
yearStart = Cells(2, 3).Value
Dim yearEnd As Double
Dim percentChange As Double


'Determine the number of rows in the sheet
Dim rowCount As Long

rowCount = Range("A1").End(xlDown).Row

Dim i As Long

'loop through all the stocks
For i = 2 To rowCount
 

    ' Check if we are still within the same stock ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
    
    ' Set the ticker name
      ticker = Cells(i, 1).Value
         
    ' make sure the first one is acounted for as well
      totalVolume = totalVolume + Cells(i, 7).Value
     
    ' Add one to the summary table row
      summaryRow = summaryRow + 1
      
    ' Set the totalVolume to the summary row
      Range("K" & summaryRow).Value = totalVolume
      Cells(i, 11).NumberFormat = "00"
      
      yearEnd = Cells(i, 6).Value
      
      yearlyChange = yearEnd - yearStart
      
    ' Set the yearlyChange to the summary row
      Range("I" & summaryRow).Value = yearlyChange

    ' Print the ticker name in the Summary Table
      Range("H" & summaryRow).Value = ticker
      
    ' find and print the percent change as a percent
      percentChange = (yearEnd - yearStart) / yearStart
      Range("J" & summaryRow).Value = percentChange
      Range("J" & summaryRow).NumberFormat = "0.00%"
      
             
    'Reset variables
    
    totalVolume = 0
    yearlyChange = 0
    percentChange = 0
    yearEnd = 0
    yearStart = Cells(i + 1, 3).Value

    Else
    
    'Add to the total volume
    totalVolume = totalVolume + Cells(i, 7).Value

    End If
    
Next i

'define more variables
Dim maxIncrease As Double
maxIncrease = 0
Dim maxDecrease As Double
maxDecrease = 0
Dim maxVolume As Double
maxVolume = 0

'row count of summary
Dim rowCountSummary As Long

rowCountSummary = Range("H1").End(xlDown).Row

'loop through summary
For n = 2 To rowCountSummary

'deermine the max values for % increase, decrease, and total volume
If Cells(n, 10).Value > maxIncrease Then
    maxIncrease = Cells(n, 10).Value
    ticker = Cells(n, 8).Value
    Cells(2, 15).Value = ticker
End If
Cells(2, 16).Value = maxIncrease
Cells(2, 16).NumberFormat = "0.00%"

If Cells(n, 10).Value < maxDecrease Then
    maxDecrease = Cells(n, 10).Value
    ticker = Cells(n, 8).Value
    Cells(3, 15).Value = ticker
End If
Cells(3, 16).Value = maxDecrease
Cells(3, 16).NumberFormat = "0.00%"

If Cells(n, 11).Value > maxVolume Then
    maxVolume = Cells(n, 11).Value
    ticker = Cells(n, 8).Value
    Cells(4, 15).Value = ticker
End If
Cells(4, 16).Value = maxVolume
Cells(4, 16).NumberFormat = "00"

'Conditional Formatting to color positive changes green and negative red
If Cells(n, 9).Value < 0 Then
    Cells(n, 9).Interior.ColorIndex = 3 'Red
    Cells(n, 10).Interior.ColorIndex = 3 'Red
ElseIf Cells(n, 9).Value > 0 Then
   Cells(n, 9).Interior.ColorIndex = 4 'green
   Cells(n, 10).Interior.ColorIndex = 4 'green
End If

Next n

'size all columns so it's easier to read
Columns("A:P").AutoFit

End Sub

