Attribute VB_Name = "Module3"
Sub Stocks():

    For Each ws In Worksheets
    
        ' This is where we define all of our variables
        
        Dim ticker_name As String
        Dim open_sum As Double
        Dim closing_sum As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim volume_Sum As Double
        Dim lastRow As Long
        Dim row_table As Integer
        Dim count As Long
        
        ' Assigning a value to our defined variables
        
        open_sum = 0
        closing_sum = 0
        yearly_change = 0
        percent_change = 0
        volume_Sum = 0
        lastRow = Cells(Rows.count, 1).End(xlUp).Row
        row_table = 2
        count = 0
        
        ' Placing the Header into the cells and autofitting it
        
        ws.Cells(1, 9).Value = "Ticket"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        
        
        
            For i = 2 To lastRow
                
                ' It will execute this for loop when it finds a different ticker name
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ' Grabbing the ticker name and placing it into the table
                    ticker_name = ws.Cells(i, 1).Value
                    ws.Range("I" & row_table).Value = ticker_name
                    
                    ' This grabs the opening number from the first day
                    open_sum = ws.Cells(i - count, 3).Value
                    ' This grabs the closing number on the last day
                    closing_sum = ws.Cells(i, 6).Value
                    ' This is a running sum for each volume cell in that ticker range
                    volume_Sum = volume_Sum + ws.Cells(i, 7).Value
                    ' Yearly change is just the closing day total - the opening day total, and places it into the table
                    yearly_change = closing_sum - open_sum
                    ws.Range("J" & row_table).Value = yearly_change
                    
                    ' This is how I accounted for dividing by 0
                    ' Then it changed the cell color, calculated the percent change, then formatted the number into a percent
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
                    ' Places the volume sum into the table
                     ws.Range("L" & row_table).Value = volume_Sum
                     
                     ' This is where we reset and increase the row count.
                     row_table = row_table + 1
                     open_sum = 0
                     closing_sum = 0
                     volume_Sum = 0
                     yearly_change = 0
                     percent_change = 0
                     count = 0
                ' Executes this else statement if its still within the same ticker
                Else
                    
                    closing_sum = Cells(i, 6).Value
                    volume_Sum = volume_Sum + Cells(i, 7).Value
                    count = count + 1
                    
                End If
            
            Next i
            
        
    Next ws


End Sub

