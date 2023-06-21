Sub homework2code()

'Set up looping through worksheets
For Each ws In Worksheets

    'name variables, assign data types and starting values, if relevant
    
    Dim i As Long
    Dim ticker_value As String
    
    Dim vol_sum As Double
    vol_sum = 0
    
    Dim yearly_change As Double
    Dim first_open As Double
    Dim last_close As Double
    Dim percent_change As Double
    
    Dim percent_change_min As Double
    percent_change_min = 0
    
    Dim percent_change_max As Double
    percent_change_max = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    Dim percent_summary As Double
    percent_summary = 0
    
    Dim vol_result As Double
    vol_result = 0
    
    'i had a hard time figuring out the max/min functions, so instead i found the greatest/smallest yearly changes by adding a pretend variable
    Dim dummy_var As Long

    'label new columns and stuff that are created
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'set length of loop to be the length of stuff
    Dim rowCount As Long
    
    
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Set up for loop for i 2 through 999999 (whatever length)
    For i = 2 To rowCount
    
        'if the cells are not equal (top bun)
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
            first_open = ws.Cells(i, 3).Value
            vol_sum = vol_sum + ws.Cells(i, 7).Value
            
        
        
        'if the cells are not equal (bottom bun, it is a hamburger)
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'calculate variables for dataset
            ticker_value = ws.Cells(i, 1).Value
            vol_sum = vol_sum + ws.Cells(i, 7).Value
            last_close = ws.Cells(i, 5).Value
            yearly_change = last_close - first_open
            percent_change = yearly_change / first_open
            
            
            'print stuff in the new columns
            ws.Range("I" & summary_table_row) = ticker_value
            ws.Range("J" & summary_table_row) = yearly_change
            
            ws.Range("K" & summary_table_row) = percent_change
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            
            ws.Range("L" & summary_table_row) = vol_sum
            
                'greatest total vol summary
                If vol_sum >= vol_result Then
                    ws.Cells(4, 17) = vol_sum
                    vol_result = vol_sum
                    ws.Cells(4, 16) = ticker_value
                Else
                    dummy_var = dummy_var + 1
                End If
                
                'yearly change greatest increase value
                If percent_change >= percent_change_max Then
                    ws.Cells(2, 17) = percent_change
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                    percent_change_max = percent_change
                    ws.Cells(2, 16) = ticker_value
                Else
                    dummy_var = dummy_var + 1
                End If
                
                'yearly change greatest decrease value
                If percent_change <= percent_change_min Then
                    ws.Cells(3, 17) = percent_change
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    percent_change_min = percent_change
                    ws.Cells(3, 16) = ticker_value
                Else
                    dummy_var = dummy_var + 1
                End If
            
                'format the yearly change column
                If yearly_change >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    
                End If
                

            'move the summary column down a row
            summary_table_row = summary_table_row + 1
            
            'reset the values for the next i
            vol_sum = 0
            yearly_change = 0
       
            last_close = 0
            first_open = 0
            
        'the hamburger itself
        Else
            vol_sum = vol_sum + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    
Next ws
    


End Sub
