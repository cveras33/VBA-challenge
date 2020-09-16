Attribute VB_Name = "Module1"
Sub VBA_Stocks():

    ' Variable to store the last row
    ' Dim last_row As Long
    
    ' Variable to store the current volume, which can either be added to or outputted per ticker
    Dim volume As Long
    
    ' Varaible to store the ticker symbol which will then be outputted
    Dim ticker As String
    
    'Variable for specifying column of interest
    Dim column As Integer
    column = 1
    
    ' Variable for location of outputs such as ticker, yearly change, % change & total stock vol.
    Dim output_row As Integer
    output_row = 2
    
    ' Determining the last row -- WASN'T WORKING PROPERLY
    ' last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Looping variable
    Dim i As Long
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Looping through every row
    For i = 2 To 70926
    
        ' Determining if the ticker has changed and if it has:
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
             
            'set the ticker to the last ticker symbol
            ticker = Cells(i, 1).Value
                 
            ' print ticker symbol into appropriate row
            Range("I" & output_row).Value = ticker
                 
            ' adding to the volume
            volume = volume + Cells(i, 7).Value
                  
            ' print volume sum into appropriate row
            Range("L" & output_row).Value = volume
                 
            ' Increment the output_row
            output_row = output_row + 1
                 
            ' Reset volume sum
            volume = 0
        
        Else
            ' Add to volume sum -- GETTING BUG HERE
            volume = volume + Cells(i, 7).Value
            ' volume = CLng(volume + Cells(i, 7).Value)
            
        End If
    
    Next i
    
    ' Formating
    ' positive / negative (from checker board example)
    ' percentages (from wells_fargo example -- similar to how currency was formatted)

End Sub
