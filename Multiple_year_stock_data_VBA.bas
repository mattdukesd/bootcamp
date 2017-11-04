Attribute VB_Name = "Module1"
Sub wall_street():

Dim ticker As String
Dim vol As Double
vol = 0
Dim result_row As Integer
results_row = 2

Dim high As Double

Dim low As Double


Dim tick_header As String
tick_header = "Ticker"
Range("I1").Value = tick_header

Dim high_header As String
high_header = "High"
Range("K1").Value = high_header

Dim low_header As String
low_header = "Low"
Range("L1").Value = low_header

Dim vol_header As String
vol_header = "Volume"
Range("J1").Value = vol_header

Dim opening_header As String
opening_header = "Open"
Range("M1").Value = opening_header

Dim opening As Double

Dim closing As Double

Dim closing_header As String
closing_header = "Close"
Range("N1") = closing_header

Dim value_change As Double

Dim value_header As String
value_header = "Change"
Range("O1").Value = value_header

Dim percent_change As Double

Dim percent_header As String
percent_header = "Percent"
Range("P1").Value = percent_header


Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    'check to see if the next value is still within the same credit card brand
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set the card brand because you know the next one will be different
        ticker = Cells(i, 1).Value
        'this is the final row for the card so add to the final row to the card total
        vol = vol + Cells(i, 7).Value
        'this is the final row for the ticker, so grab the closing price
        closing = Cells(i, 6).Value
        'the next line is the opening value for the next ticker, so grab that value
        opening = Cells(i + 1, 6).Value
        'print the card brand in the summary table
        Range("I" & results_row).Value = ticker
        'print the amount to the table
        Range("J" & results_row).Value = vol
        'print the high
        Range("K" & results_row).Value = high
        'print the low
        Range("L" & results_row).Value = low
        'print closing value
        Range("N" & results_row).Value = closing
        'print opening value
        Range("M" & results_row + 1).Value = opening
        
        
        
        
        
        'add one to the table row so that the next card will go down one row
        results_row = results_row + 1
        'reset the brand total because the next line is starting a different card
        vol = 0
        
        'set the value of high so that the first row of the next ticker is included in the else formula
        high = Cells(i + 1, 4).Value
        'set the value of low so that the first row of the next ticker is included in the below else formula
        low = Cells(i + 1, 5).Value
        
    'If the cell following a row is the same brand
    Else
        'add to the running total
        vol = vol + Cells(i, 7).Value
        
        'Is this row greater than the row below?
        
        
        If Cells(i, 4) > Cells(i + 1, 4) Then
            'Keep the value if it's greater
            high = Cells(i, 4).Value
        Else
            'Keep the second value if greater
            high = Cells(i + 1, 4).Value
        End If
        
        
                'Is this row less than the row below?
        If Cells(i, 5) < Cells(i + 1, 5) Then
            'Keep the value if it's lower
            low = Cells(i, 5).Value
        Else
            'Keep the second value if it's lower
            low = Cells(i + 1, 5).Value
        End If
        
        

    End If
    
Next i


For i = 2 To LastRow

'Subtract the close from the open to determine the change over the year
value_change = ((Cells(i, 14).Value) - (Cells(i, 13).Value))
'Print the change
Cells(i, 15).Value = value_change

    'Create if statement to change cell color
    If (Cells(i, 15) > 0) Then
        'Green if positive
        Cells(i, 15).Interior.ColorIndex = 4
        
    Else
        'Red if negative
        Cells(i, 15).Interior.ColorIndex = 3
        
    End If
    
    
    

Next i

End Sub



