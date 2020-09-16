Attribute VB_Name = "Module1"
Sub VBA_Stocks():

    ' Declare Variables
    Dim last_row As Long
    Dim curr_volume As Long
    Dim new_volume As Long
    Dim ticker As String
    
    'Variable for specifying column of interest
    Dim column As Integer
    column = 1
    
    ' Variable for location of outputs
    Dim outout_row As Integer
    output_row = 2
    
    ' Find last row
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Looping through the every row
    For i = 2 To last_row
    
        ' Need to determine where tickers change (did this in class - NextCell example) --> if else statement
        ' If the ticker change then --
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
             
             ' Make a record of each ticker when the ticker changes --
             ticker = Cells(i, 1).Value
             
             ' Must also print and keep track of row the ticker is being recorded in (incrementing)
             Range("I" & output_row).Value = ticker
             
        ' Else
        Else
            ' sum volume
            
        End If
    
    Next i
    
    ' Formating
    ' positive / negative (from checker board example)
    ' percentages (from wells_fargo example -- similar to how currency was formatted)

End Sub
