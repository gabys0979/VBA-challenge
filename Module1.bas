Attribute VB_Name = "Module1"
Sub yearlystocks()
'set variables
Dim ticker As String
Dim Openprice As Double
Dim Closeprice As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim TotalVolume As Double
Dim Summary_Table As Integer

Openprice = Cells(2, 3).Value
Summary_Table = 2

TotalVolume = 0


'ticker symbol
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount


If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       ticker = Cells(i, 1).Value
       
        Closeprice = Cells(i, 6).Value
        'yearlychange last price of the year - the open price of the year
        
        
        Yearly_Change = Closeprice - Openprice
        
  'percent change = yearly change/openprice of the year
  
       Percent_Change = Yearly_Change / Openprice
       
       TotalVolume = TotalVolume + Cells(i, 7).Value
        Openprice = Cells(i + 1, 3).Value
        
'Headers
Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly_Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Volume"



        
      'summary table I=ticker symbol, J= Yearly change
      
        Range("I" & Summary_Table).Value = ticker
        Range("J" & Summary_Table).Value = Yearly_Change
        Range("k" & Summary_Table).Value = FormatPercent(Percent_Change)
        Range("L" & Summary_Table).Value = TotalVolume
        Summary_Table = Summary_Table + 1
        
        
TotalVolume = 0

'if same keep totaling
Else


    TotalVolume = TotalVolume + Cells(i, 7).Value
    'keep getting error
    
  
 
 End If
 
 Next i
 
'bonus headers

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'Bonus description

Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest total volume"

'Greatest % increase
Range("P2").Value = FormatPercent(WorksheetFunction.Max(Range("k2:k" & RowCount)))

'greatest decrease

Range("P3").Value = FormatPercent(WorksheetFunction.Min(Range("k2:k" & RowCount)))

'greatest volume

Range("P4").Value = WorksheetFunction.Max(Range("l2:l" & RowCount))

'Range("S2:S").Value = ticker

End Sub

