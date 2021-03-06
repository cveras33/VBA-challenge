Attribute VB_Name = "Module1"
Sub VBA_Stocks():
    
    For Each ws In Worksheets
    
        ' Variable to store the last row
        Dim last_row As Long
        
        ' Getting the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Varaible to store the ticker symbol
        Dim ticker As String
        
        ' Variable to store yearly change
        Dim yearly_change As Double
        yearly_change = 0
        
        'Variable to store percent change
        Dim percent_change As Double
               
        ' Variable to store the volume sum
        Dim volume As Double
        volume = 0
        
        ' Variable to store opening value, and setting that value to 0
        Dim opening_value As Double
        opening_value = 0
        
        ' Variable to store closing value, and setting that value to 0
        Dim closing_value As Double
        closing_value = 0
        
        'Variable for specifying column of interest
        Dim column As Integer
        column = 1
        
        ' Variable for starting location of outputs
        Dim output_row As Integer
        output_row = 2
        
        ' CHALLENGE Variables
        ' Variable for Greatest % increase
        Dim percent_increase As Double
        percent_increase = 0
        
        ' Variable for the ticker with the greatest percent increase
        Dim increase_ticker As String
        
        ' Variable for Greatest % decrease
        Dim percent_decrease As Double
        percent_decrease = 0
        
        ' Variable for the ticker with the greatest percent decrease
        Dim decrease_ticker As String
        
        ' Variable for Greatest Total Volume
        Dim greatest_total_volume As Double
        greatest_total_volume = 0
        
        ' Variable for the ticker with greatest total volume
        Dim volume_ticker As String
        
        ' Looping variable
        Dim i As Long
        
        ' Print Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Print CHALLENGE Headers
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Looping through every row
        For i = 2 To last_row
        
            ' Get opening value
            If opening_value = 0 Then
                
                opening_value = ws.Cells(i, 3).Value
                
            End If
        
            ' Determining if the ticker has changed and if it has:
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                 
                ' Set the ticker
                ticker = ws.Cells(i, 1).Value
                     
                ' Print ticker symbol into appropriate row
                ws.Range("I" & output_row).Value = ticker
                     
                ' Add to the volume
                volume = volume + ws.Cells(i, 7).Value
                      
                ' Print volume sum
                ws.Range("L" & output_row).Value = volume
                
                ' Get closing price
                closing_value = ws.Cells(i, 6).Value
                
                ' Calculating Yearly Change
                yearly_change = closing_value - opening_value
                
                ' Print yearly change
                ws.Range("J" & output_row).Value = yearly_change
                
                ' Formatting yearly change
                ' Positive values will be filled green
                If yearly_change > 0 Then
                
                    ws.Range("J" & output_row).Interior.ColorIndex = 4
                
                ' Negative values will be filled red
                Else
                
                    ws.Range("J" & output_row).Interior.ColorIndex = 3
                
                End If
                
                If opening_value <> 0 Then
                    ' Calculate percent change
                    percent_change = (yearly_change / Abs(opening_value))
                Else
                
                    percent_change = 0
                
                End If
                
                ' Print percent change
                ws.Range("K" & output_row).Value = percent_change
                
                ' Formatting percent change to a percentage
                ws.Range("K" & output_row).Style = "Percent"
                
                ' Formatting percentage to go to 2 decimal places
                ws.Range("K" & output_row).NumberFormat = "0.00%"
                
                ' Resetting for the next ticker
                ' Increment output_row
                output_row = output_row + 1
                     
                ' Reset volume sum
                volume = 0
                
                opening_value = 0
            
            Else
                ' Add to volume sum
                volume = volume + ws.Cells(i, 7).Value
                
                ' CHALLENGE
                ' Determining the greatest total volume
                If volume > greatest_total_volume Then
                    
                    greatest_total_volume = volume
                    volume_ticker = ws.Cells(i, 1).Value
                
                End If
                
                ' Determining the greatest percent increase
                If percent_change > percent_increase Then
                
                    percent_increase = percent_change
                    increase_ticker = ws.Cells(i, 1).Value
                
                End If
                
                'Determining the greatest percent decrease
                If percent_change < percent_decrease Then
                
                    percent_decrease = percent_change
                    decrease_ticker = ws.Cells(i, 1).Value
                
                End If
                
            End If
        
        Next i
        
        
        ' Print greatest total volume and its ticker
        ws.Range("P4").Value = volume_ticker
        ws.Range("Q4").Value = greatest_total_volume
        
       ' Print greatest percent decrease and its ticker
        ws.Range("Q3").Value = percent_decrease
        ws.Range("P3").Value = decrease_ticker
        
        ' Formatting to a percentage
        ws.Range("Q3").Style = "Percent"
                
        ' Formatting percentage to go to 2 decimal places
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ' Print greatest percent increase and its ticker
        ws.Range("Q2").Value = percent_increase
        ws.Range("P2").Value = increase_ticker
        
        ' Formatting to a percentage
        ws.Range("Q2").Style = "Percent"
                
        ' Formatting percentage to go to 2 decimal places
        ws.Range("Q2").NumberFormat = "0.00%"

    Next ws

End Sub

