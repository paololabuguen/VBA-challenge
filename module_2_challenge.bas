Attribute VB_Name = "Module1"
Sub module_2_challenge()

    'Loop through worksheets
    For Each cur_sheet In Worksheets
        'Create Headers on right table
        cur_sheet.Range("I1").Value = "Ticker"
        cur_sheet.Range("J1").Value = "Yearly Change"
        cur_sheet.Range("K1").Value = "Percent Change"
        cur_sheet.Range("L1").Value = "Total Stock Volume"
        
        'Create the row and column headers for the greatest increase, decrease and volume
        cur_sheet.Range("Q1").Value = "Ticker"
        cur_sheet.Range("R1").Value = "Value"
        cur_sheet.Range("P2").Value = "Greatest % Increase"
        cur_sheet.Range("P3").Value = "Greatest % Decrease"
        cur_sheet.Range("P4").Value = "Greatest Total Volume"
    
        'Define Variables
        Dim ticker As String
    
        'Loop Variables
        Dim t_check As String
        Dim ind1 As Long  'Index for the main while loop
        Dim ind2 As Long  'This index is for the table on the right when we input
                          'values on the cells
    
        ind1 = 2 'Initial value for the data index
    
        ind2 = 2 'Initial value for the table index
    
        'Variables for Prices
        Dim op_price As Double  'Opening price
        Dim cl_price As Double  'Closing price
        Dim volume As Double    'Volume
    
        While IsEmpty(cur_sheet.Cells(ind1, 1)) = False
    
            'Reset the values for the next ticker
            t_check = cur_sheet.Cells(ind1, 1).Value
            ticker = cur_sheet.Cells(ind1, 1).Value
            volume = 0
            op_price = cur_sheet.Cells(ind1, 3).Value
        
            'Checking to see if we are at the end of the year row
            While t_check = ticker
                volume = volume + cur_sheet.Cells(ind1, 7).Value
                ind1 = ind1 + 1
                t_check = cur_sheet.Cells(ind1, 1).Value
            Wend
        
            'Retrieve the closing price when we reach end of the year
            cl_price = cur_sheet.Cells(ind1 - 1, 6).Value
        
            'Input the appropriate values on the table on the right
            cur_sheet.Cells(ind2, 9).Value = ticker                    'Ticker Value
            cur_sheet.Cells(ind2, 10) = cl_price - op_price            'Yearly Change Value
            cur_sheet.Cells(ind2, 11) = cur_sheet.Cells(ind2, 10) / op_price  'Percentage Change Value
            cur_sheet.Cells(ind2, 12) = volume                         'Total Volume
        
            'Check if the new percent change is larger than the previous largest
            If (cur_sheet.Cells(ind2, 10) / op_price) > cur_sheet.Range("R2").Value Then
        
                'Replace new value if it is
                cur_sheet.Range("R2").Value = cur_sheet.Cells(ind2, 10) / op_price
                cur_sheet.Range("Q2").Value = ticker
            End If
        
            'Check if the new percent change is smaller than the previous smallest
            If (cur_sheet.Cells(ind2, 10) / op_price) < cur_sheet.Range("R3").Value Then
        
                'Replace new value if it is
                cur_sheet.Range("R3").Value = cur_sheet.Cells(ind2, 10) / op_price
                cur_sheet.Range("Q3").Value = ticker
            End If
        
            'Check if the new volume is larger than the previous largest
            If cur_sheet.Cells(ind2, 12) > cur_sheet.Range("R4").Value Then
        
                'Replace new value if it is
                cur_sheet.Range("R4").Value = cur_sheet.Cells(ind2, 12)
                cur_sheet.Range("Q4").Value = ticker
            End If
        
        
            'Formatting the table on the right
            cur_sheet.Cells(ind2, 11).Value = FormatPercent(cur_sheet.Cells(ind2, 11))
        
            'Format for the greatest change table
            cur_sheet.Range("R2").Value = FormatPercent(cur_sheet.Range("R2"))
            cur_sheet.Range("R3").Value = FormatPercent(cur_sheet.Range("R3"))
        
            If cur_sheet.Cells(ind2, 10) < 0 Then
                'Color the cell red if Yearly Change > 0
                cur_sheet.Cells(ind2, 10).Interior.ColorIndex = 3
            
            ElseIf cur_sheet.Cells(ind2, 10) > 0 Then
                'Color the cell green if Yearly Change < 0
                cur_sheet.Cells(ind2, 10).Interior.ColorIndex = 4
            
            Else
                'Color the cell gray if Yearly Change = 0
                cur_sheet.Cells(ind2, 11).Interior.ColorIndex = 16
            End If

            ind2 = ind2 + 1 'Next index for the subsheet row
        Wend
    
    Next cur_sheet
    
End Sub
