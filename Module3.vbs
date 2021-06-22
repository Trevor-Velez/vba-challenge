Attribute VB_Name = "Module3"
Sub Stocks():

    For Each ws In Worksheets
    
    
        Dim ticker_name As String
        Dim open_sum As Double
        Dim closing_sum As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim volume_Sum As Double
        Dim lastRow As Long
        Dim row_table As Integer
        Dim count As Long
        
        open_sum = 0
        closing_sum = 0
        yearly_change = 0
        percent_change = 0
        volume_Sum = 0
        lastRow = Cells(Rows.count, 1).End(xlUp).Row
        row_table = 2
        count = 0
        
        ws.Cells(1, 9).Value = "Ticket"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        
        
        
            For i = 2 To lastRow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ticker_name = ws.Cells(i, 1).Value
                    ws.Range("I" & row_table).Value = ticker_name
                    
                    open_sum = ws.Cells(i - count, 3).Value
                    closing_sum = ws.Cells(i, 6).Value
                    volume_Sum = volume_Sum + ws.Cells(i, 7).Value
                    
                    yearly_change = closing_sum - open_sum
                    ws.Range("J" & row_table).Value = yearly_change
                    
                    
                        If yearly_change > 0 Then
                            If open_sum = 0 Then
                                ws.Range("J" & row_table).Interior.ColorIndex = 4
                                percent_change = ((closing_sum - open_sum))
                                ws.Range("K" & row_table).Value = percent_change
                                ws.Range("K" & row_table).NumberFormat = "0.00%"
                        
                            Else
                                ws.Range("J" & row_table).Interior.ColorIndex = 4
                                percent_change = ((closing_sum - open_sum) / open_sum)
                                ws.Range("K" & row_table).Value = percent_change
                                ws.Range("K" & row_table).NumberFormat = "0.00%"
                            End If
                                                
                        Else
                            If open_sum = 0 Then
                                ws.Range("J" & row_table).Interior.ColorIndex = 3
                                percent_change = ((open_sum - closing_sum))
                                ws.Range("K" & row_table).Value = "-" & percent_change
                                ws.Range("K" & row_table).NumberFormat = "0.00%"
                            
                            Else
                                ws.Range("J" & row_table).Interior.ColorIndex = 3
                                percent_change = ((open_sum - closing_sum) / open_sum)
                                ws.Range("K" & row_table).Value = "-" & percent_change
                                ws.Range("K" & row_table).NumberFormat = "0.00%"
                            End If
                        End If
                      
                     ws.Range("L" & row_table).Value = volume_Sum
                     
                     row_table = row_table + 1
                     open_sum = 0
                     closing_sum = 0
                     volume_Sum = 0
                     yearly_change = 0
                     percent_change = 0
                     count = 0
                    
                Else
                    
                    closing_sum = Cells(i, 6).Value
                    volume_Sum = volume_Sum + Cells(i, 7).Value
                    count = count + 1
                    
                End If
            
            Next i
            
        
    Next ws


End Sub

