Sub WorksheetLoop()

Dim ws As Worksheet

'Loop through each worksheet
For Each ws In Worksheets

    'Place titles for summary table
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    
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
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Set Inital Summary Table Row
    Dim ST_Row As Long
        ST_Row = 2
    
    Open_Price = ws.Cells(2, 3).Value
    Tot_Volume = 0
        
    'Write For Statement
    For i = 2 To (LastRow - 1)
        
        'If statement to find last row of each ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'Set end values
            End_Price = ws.Cells(i, 6).Value
            End_Volume = ws.Cells(i, 7).Value
            
            'Find Yearly Change
            Year_Change = End_Price - Open_Price
            
            'If statement to ignore open price of zero
            If Open_Price = 0 Then
            
                ws.Cells(i, 3).Value = Null
                
            Else
                
                'Find Percent Change
                Percent_Change = Year_Change / Open_Price
                ws.Cells(ST_Row, 11).Value = Percent_Change
            
            End If
            
            'Add final stock volume value
            Tot_Volume = Tot_Volume + ws.Cells(i, 7).Value
            
            'Print Values
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(ST_Row, 9).Value = Ticker
            ws.Cells(ST_Row, 10).Value = Year_Change
            ws.Cells(ST_Row, 12).Value = Tot_Volume
            
            'Conditional Formating
            ws.Cells(ST_Row, 11).NumberFormat = "0.00%"
            
            If ws.Cells(ST_Row, 10).Value > 0 Then
                ws.Cells(ST_Row, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(ST_Row, 10).Value < 0 Then
                ws.Cells(ST_Row, 10).Interior.ColorIndex = 3
            
            End If
            
            'Reset Values
            Open_Price = ws.Cells(i + 1, 3).Value
            Tot_Volume = 0
            
            'Go to next row of summary table
            ST_Row = ST_Row + 1
            
        Else
            
            'Tally total volume
            Tot_Volume = Tot_Volume + ws.Cells(i, 7).Value
            
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
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Find greatest % increase and ticker
    Increase = WorksheetFunction.Max(ws.Range("K:K"))
    TR_Increase = WorksheetFunction.Match(Increase, ws.Range("K:K"), 0)
    T_Increase = ws.Cells(TR_Increase, 9).Value
    
    'Find greatest % decrease and ticker
    Decrease = WorksheetFunction.Min(ws.Range("K:K"))
    TR_Decrease = WorksheetFunction.Match(Decrease, ws.Range("K:K"), 0)
    T_Decrease = ws.Cells(TR_Decrease, 9).Value
    
    'Find total volume and ticker
    Volume = WorksheetFunction.Max(ws.Range("L:L"))
    TR_Volume = WorksheetFunction.Match(Volume, ws.Range("L:L"), 0)
    T_Volume = ws.Cells(TR_Volume, 9).Value
    
    'Format needed cells to percentage
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Print values in cells
    ws.Range("Q2").Value = Increase
    ws.Range("P2").Value = T_Increase
    
    ws.Range("Q3").Value = Decrease
    ws.Range("P3").Value = T_Decrease
    
    ws.Range("Q4").Value = Volume
    ws.Range("P4").Value = T_Volume
    
    ws.Activate
    
Next
    
End Sub
