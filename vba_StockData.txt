Sub ticker_count()
    'Declare Variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume_Total As Double
    Dim year_Open As Double
    Dim year_Close As Double
    Dim yearly_Change As Double
    Dim percent_Change As Double

        'To declare and activate each worksheet in workbook
        
        For Each ws In Worksheets
        ws.Activate
        ws.Cells.EntireColumn.AutoFit
        
        
        'Set inital volume to 0
        
        'Set inital display column number
        
        'Find last row in of data in each worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Insert the header to each worksheets for column "I" & "J"
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set inital volume and first Year Open
        volume = 0
        year_Open = Cells(2, 3).Value
        summary_Display = 2
                
                'Iterate from row 2 to last row
                For i = 2 To lastRow
                    
                    'Check the to see if the next ticker is the same as the last ticker
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                        ticker = Cells(i, 1).Value
                        volume = volume + Cells(i, 7).Value
                        year_Close = Cells(i, 6).Value
                        
                        'See if the next ticker is starting with year_Open = 0 and year_Close <> 0, then iterate to till year_Open does not equal 0 starting from the Count.
                        
                        If (year_Open = 0 And year_Close <> 0) Then
                              Do While Cells(Count, 3).Value = 0
                                ' The counter to find the year_Open number that does not equal to 0
                                 Count = Count + 1
                                   Loop
                                    year_Open = Cells(Count, 3).Value
                                    percent_Change = (year_Close / year_Open) - 1
                            Else
                        End If
                       
                            
                        ' Check to see if year_Open = 0 and year_Close = 0 then set percent_Change to 0
                        If (year_Open = 0 And year_Close = 0) Then
                            percent_Change = 0
                            
                            'Check if year_One is 1 and year_Close = 0 then percent_Change to 1
                            ElseIf (year_Open = 1 And year_Close <> 0) Then
                            percent_Change = 1

                            ' Calculate percent_Change    
                            Else
                                percent_Change = (year_Close / year_Open) - 1

                        End If
                            
                            'Set interior colors to Green for positive and Red for negative for yearly_Change
                            yearly_Change = year_Close - year_Open
                                If (yearly_Change >= 0) Then
                                    ws.Range("J" & summary_Display).Interior.ColorIndex = 4
                                    Else
                                    ws.Range("J" & summary_Display).Interior.ColorIndex = 3
                                End If
                            
                        
                        'Set values corresponding to ticker
                        ws.Range("I" & summary_Display).Value = ticker
                        ws.Range("L" & summary_Display).Value = volume
                        ws.Range("J" & summary_Display).Value = yearly_Change
                        ws.Range("K" & summary_Display).Value = percent_Change
                        ws.Range("K" & summary_Display).NumberFormat = "0.00%"
                        
                        
                        'Reset volume and year_open
                        volume = 0
                        year_Open = 0
                        
                        'Set year_Open for next ticker
                        year_Open = year_Open + Cells(i + 1, 3).Value
                        
                        'Set counter for year_Open
                        Count = i + 1
                        
                        'Reset year_Close
                        year_Close = 0
            
                    Else
                        'Set the new starting volume number for next ticker
                        volume = volume + Cells(i, 7).Value
                        
                    End If
                
                Next i
  
    
        
    

        'Set headers for displaying calculated values
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("p1").Value = "Value"

        'Set last row for showing ticker    
        change_LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
            'Iterate from column k to find max number and display max number with corresponding ticker and percent_Change
            For j = 2 To change_LastRow
                If Cells(j, "K").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & change_LastRow)) Then
                    Cells(2, "O").Value = Cells(j, "I").Value
                    Cells(2, "P").Value = Cells(j, "K").Value
                    Cells(2, "P").NumberFormat = "0.00%"
                
                    'Iterate from column k to find max number and display min number with corresponding ticker and percent_Change
                    ElseIf Cells(j, "K").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & change_LastRow)) Then
                    Cells(3, "O").Value = Cells(j, "I").Value
                    Cells(3, "P").Value = Cells(j, "K").Value
                    Cells(3, "P").NumberFormat = "0.00%"
                    Else
                End If
                
                'Iterate to find the max volume_Total and display with corresponding ticker
                If Cells(j, "L").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & change_LastRow)) Then
                    Cells(4, "O").Value = Cells(j, "I").Value
                    Cells(4, "P").Value = Cells(j, "L").Value
                    Else
                        
                End If
      
        
      
            Next j
  
    Next ws
        
    

End Sub