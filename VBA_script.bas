Attribute VB_Name = "Module2"
Sub stocks():

    'loop through worksheets
    For Each ws In Worksheets
    
   'labels Part 1
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'labels Part 2
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'labels for the Bonus
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
 
    'declare variables Part 1
    Dim Ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim vol As Double
        
    'declare variables Part 1
    Dim price_open As Double
    Dim price_close As Double
  
   
    Dim row As Integer
    Dim col As Integer

    Dim greatestincrease As Double
    Dim greatestdecrease As Double
       
    'initialize variables Part 1
    'summary table row
    summary_row = 2
   'total volume
    vol = 0
    
    'initialize variables Part 2
    'price initial
    price_open = Cells(2, 3).Value
    
    'initialize variables for the Bonus
    'greatest % increase
    ws.Range("Q2").Value = 0
    'greatest % decrease
     ws.Range("Q3").Value = 0
    'greatest total volume
     ws.Range("Q4").Value = 0
     
        
    'Determine last row
    row_final = ws.Cells(Rows.Count, 1).End(xlUp).row

    'loop through each ticker
    For i = 2 To row_final

        'if current ticker is different than the next ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'input current ticker name before it changes to a different one
            Ticker = ws.Cells(i, 1).Value
            'add last ticker volume in the set
            vol = vol + Cells(i, 7).Value
            
            'Calculate yearly change from opening price at beginning to closing price at the end
            price_close = ws.Cells(i, 6).Value
            yearly_change = price_close - price_open
            
            'Calculate change percent and account for zero value
            If price_open <> 0 Then
                percent_change = (yearly_change / price_open) * 100
            End If
                  
            'print out ticker name
            ws.Range("I" & summary_row).Value = Ticker
            'print out yearly change
            ws.Range("J" & summary_row).Value = yearly_change
            
            'format percent change to include %
            ws.Range("K" & summary_row).NumberFormat = "0.00\%"
            'print out percent change
            ws.Range("K" & summary_row).Value = percent_change
            'print out final volume
            ws.Range("L" & summary_row).Value = vol
            
            If (yearly_change > 0) Then
            ws.Range("J" & summary_row).Interior.ColorIndex = 4
            
            ElseIf (yearly_change <= 0) Then
            ws.Range("J" & summary_row).Interior.ColorIndex = 3
            
            End If
                                      
            summary_row = summary_row + 1

            vol = 0
            price_open = ws.Cells(i + 1, 6).Value
        
        'if current ticker is the same as the next ticker
        Else
            'Add ticker volume total
            vol = vol + ws.Cells(i, 7).Value

    End If
    
Next i

'Greatest Percent Increase, Greatest Percent Decrease, and Total Volume #bonus

    For j = 2 To row_final
    'checks for greatest % increase
        If ws.Range("K" & j).Value > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Range("K" & j).Value
        ws.Range("P2").Value = ws.Range("I" & j).Value
        End If
    
    'checks for greatest % decrease
        If ws.Range("K" & j).Value < ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Range("K" & j).Value
        ws.Range("P3").Value = ws.Range("I" & j).Value
        End If
        
    'checks for greatest total stock volume
        If ws.Range("L" & j).Value > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Range("L" & j).Value
        ws.Range("P4").Value = ws.Range("I" & j).Value
        End If
            
     ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit
        ws.Range("Q2:Q3").NumberFormat = "0.00\%"
      
    Next j
    

Next ws
 
End Sub


