Sub Analysis2()
 
 'Starting the For each loop within all worksheets
For Each ws In Worksheets
    
    'Declaring Variables to be used in the macro
    Dim Ticker As String
    Dim Yearly, Percent, Summary As Double
    Dim LastRow, Volume As Double
    Dim Opening, Closing As Double
    
    'Setting up intial variable sum
    Volume = 0
    Summary = 2
    Opening = 0
    Yearly = 0
    Percent = 0
    
        'Getting the firt row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    'Setting headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

   'Skipping the header and looping to find tickers and other info
    For i = 2 To LastRow
        
       Ticker = ws.Cells(i, 1).Value
       Volume = Volume + ws.Cells(i, 7).Value
       
       'Finding the starting price of the ticker open
       If Opening = 0 Then
        Opening = ws.Cells(i, 3).Value
        End If
        
        'Find ticker value if they are not the same from prior row
        If ws.Cells(i + 1, 1).Value <> Ticker Then
        
        ws.Range("I" & Summary).Value = Ticker
        
        Closing = ws.Cells(i, 6).Value
        
        'Finding the price change value
        Yearly = Closing - Opening
        
        ws.Range("J" & Summary).Value = Yearly
       
         ws.Range("L" & Summary).Value = Volume
         Volume = 0
        'Conditionally formatting the stock value change
        If Yearly > 0 Then
            ws.Range("J" & Summary).Interior.ColorIndex = 4
        ElseIf Yearly < 0 Then
            ws.Range("J" & Summary).Interior.ColorIndex = 3
        Else
        ws.Range("J" & Summary).Interior.ColorIndex = 44
       
        End If
        
        'Finding the percent change
        If Opening = 0 Then
            Percent = 0
        Else
            Percent = Yearly / Opening
        End If
        
        ws.Range("K" & Summary).Value = Format(Percent, "Percent")
            
            'Conditionally formatting the percent values
            If Percent > 0 Then
                ws.Range("K" & Summary).Interior.ColorIndex = 4
            ElseIf Percent < 0 Then
              ws.Range("K" & Summary).Interior.ColorIndex = 3
              Else
                ws.Range("K" & Summary).Interior.ColorIndex = 44
            
             End If
            'Resetting opening price to get a different ticker
            Opening = 0
        'Adding to the row to find the next value for each subsquent table values
        Summary = Summary + 1
        End If
    Next i
  'Decrlaring variables for greatest values
  Dim GrPerInc, GrPerDec, GrVol As Double
  Dim LastRowB
  'Refinding the lastrow with a different style
  LastRowB = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  'Using formulas to find the greatest values to per change, per dec and volume
  GrPerInc = WorksheetFunction.Max(Range(ws.Cells(2, 11), ws.Cells(LastRowB, 11)))
    GrPerDec = WorksheetFunction.Min(Range(ws.Cells(2, 11), ws.Cells(LastRowB, 11)))
    GrVol = WorksheetFunction.Max(Range(ws.Cells(2, 12), ws.Cells(LastRowB, 12)))
  
  

    'Setting the loop to find the vlaues
    For j = 2 To LastRowB
        
        'If statement the dervies each values
        If ws.Cells(j, 11).Value = GrPerInc Then
        ws.Cells(2, 16).Value = Cells(j, 9).Value
        
        
        ElseIf ws.Cells(j, 11).Value = GrPerDec Then
        ws.Cells(3, 16).Value = Cells(j, 9).Value
     
        ElseIf ws.Cells(j, 12).Value = GrVol Then
        ws.Cells(4, 16).Value = Cells(j, 9).Value
        
    
    End If
    Next j
    
    'formatting to percent and setting volume to summary table
    ws.Range("Q2").Value = Format(GrPerInc, "Percent")
    ws.Range("Q3").Value = Format(GrPerDec, "Percent")
    ws.Range("Q4").Value = GrVol

    Next ws


End Sub
