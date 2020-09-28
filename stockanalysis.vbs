Sub StockAnalysis()

'Place titles for summary table
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Stock Volume"

'ID Values
Dim Ticker As String
Dim Open_Price As Double
Dim End_Price As Double
Dim Year_Change As Double
Dim Percent_Change As Double
Dim Initial_Volume As Long
Dim Tot_Volume As Double

'Find Last Row
Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Set Inital Summary Table Row
Dim ST_Row As Long
    ST_Row = 2

Open_Price = Cells(2, 3).Value
Tot_Volume = 0
    
'Write For Statement
For i = 2 To (LastRow - 1)
    
    'If statement to find last row of each ticker
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        'Set end values
        End_Price = Cells(i, 6).Value
        End_Volume = Cells(i, 7).Value
        
        'Find Yearly Change
        Year_Change = End_Price - Open_Price
        
        'If statement to ignore open price of zero
        If Open_Price = 0 Then
        
            Cells(i, 3).Value = Null
            
        Else
            
            'Find Percent Change
            Percent_Change = Year_Change / Open_Price
            Cells(ST_Row, 11).Value = Percent_Change
        
        End If
        
        'Add final stock volume value
        Tot_Volume = Tot_Volume + Cells(i, 7).Value
        
        'Print Values
        Ticker = Cells(i, 1).Value
        Cells(ST_Row, 9).Value = Ticker
        Cells(ST_Row, 10).Value = Year_Change
        Cells(ST_Row, 12).Value = Tot_Volume
        
        'Conditional Formating
        Cells(ST_Row, 11).NumberFormat = "0.00%"
        
        If Cells(ST_Row, 10).Value > 0 Then
            Cells(ST_Row, 10).Interior.ColorIndex = 4
        
        ElseIf Cells(ST_Row, 10).Value < 0 Then
            Cells(ST_Row, 10).Interior.ColorIndex = 3
        
        End If
        
        'Reset Values
        Open_Price = Cells(i + 1, 3).Value
        Tot_Volume = 0
        
        'Go to next row of summary table
        ST_Row = ST_Row + 1
        
    Else
        
        'Tally total volume
        Tot_Volume = Tot_Volume + Cells(i, 7).Value
        
    End If
    
Next i

'BONUS

'ID Varibles
Dim Increase As Double
Dim Decrease As Double
Dim Volume As Double
Dim TR_Increase As Long
Dim T_Increase As String
Dim TR_Decrease As Long
Dim T_Decrease As String
Dim TR_Volume As Long
Dim T_Volume As String

'Place titles in table
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Find greatest % increase and ticker
Increase = WorksheetFunction.Max(Range("K:K"))
TR_Increase = WorksheetFunction.Match(Increase, Range("K:K"), 0)
T_Increase = Cells(TR_Increase, 9).Value

'Find greatest % decrease and ticker
Decrease = WorksheetFunction.Min(Range("K:K"))
TR_Decrease = WorksheetFunction.Match(Decrease, Range("K:K"), 0)
T_Decrease = Cells(TR_Decrease, 9).Value

'Find total volume and ticker
Volume = WorksheetFunction.Max(Range("L:L"))
TR_Volume = WorksheetFunction.Match(Volume, Range("L:L"), 0)
T_Volume = Cells(TR_Volume, 9).Value

'Format needed cells to percentage
Range("Q2:Q3").NumberFormat = "0.00%"

'Print values in cells
Range("Q2").Value = Increase
Range("P2").Value = T_Increase

Range("Q3").Value = Decrease
Range("P3").Value = T_Decrease

Range("Q4").Value = Volume
Range("P4").Value = T_Volume

End Sub

