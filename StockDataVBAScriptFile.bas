Attribute VB_Name = "Module1"
Sub VBA_Challenge()
    
    For Each ws In Worksheets
        
        ' Keep track of each ticker in the Summary Table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        ' Set an initial variable
        Dim Ticker As String
        Dim total_stock_volume As Double
        total_stock_vloume = 0
        
        'Setting variables for Yearly Changes
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim yearly_change As Double
        yearly_change = 0
        Dim change_percentage As Double
        change_percentage = 0
        
        ' Set last_row variable for holding number of last row in each worksheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set Header for each worksheet
        ws.Range("J" & 1).Value = "Ticker"
        ws.Range("K" & 1).Value = "Yearly Change"
        ws.Range("L" & 1).Value = "Percentage Change"
        ws.Range("M" & 1).Value = "Total Stock Volume"
        
        
        ' Stating if the open price has been captured
        Dim OpenPriceCaptured As Boolean
        
            ' Loop through each row of raw data
            For i = 2 To last_row
               
                If OpenPriceCaptured = False Then
            
                    ' Set opening price
                    open_price = ws.Cells(i, 3).Value
                    OpenPriceCaptured = True
                
                End If
        
                ' Check if we are still within the same Ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    ' Set Ticker
                    Ticker = ws.Cells(i, 1).Value
            
                    ' Add to total stock volume
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
                    ' Set Close Price
                    close_price = ws.Cells(i, 6).Value
            
                    ' Set yearly change
                    yearly_change = (close_price - open_price)
                    
                    ' Set if statement to avoid divison by zero
                    If open_price <> 0 Then
                        
                        ' Set change percentage for the year
                        change_percentage = (yearly_change / open_price)
                        
                    Else
                        change_percentage = 0
                    
                    End If
                    
                    ' Print ticker into Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = Ticker
            
                    ' Print yearly change into Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = yearly_change
            
                    ' Print percentage change into Summary Table and format
                    ws.Range("L" & Summary_Table_Row).Value = change_percentage
                    ws.Range("L" & Summary_Table_Row).Style = "Percent"
            
                    ' Print total stock volume into Summary Table
                    ws.Range("M" & Summary_Table_Row).Value = total_stock_volume
                    
                    ' conditional formatting of yearly changes to red(3) or green(4) or blue(5)
                    If ws.Range("K" & Summary_Table_Row).Value < 0 Then
                        
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    ElseIf ws.Range("K" & Summary_Table_Row).Value = 0 Then
                        
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 5
                    
                    Else
                        
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    End If
                    
                    ' Add one to the Summary Table Row
                    Summary_Table_Row = Summary_Table_Row + 1
            
                    ' Reset Total Volume and close price for next Ticker
                    total_stock_volume = 0
                    close_price = 0
                    yearly_change = 0
                    change_percentage = 0
                    
                    ' Set to false for the opening price of next ticker
                    OpenPriceCaptured = False
            
                Else
            
                    ' Add to total stock volume
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
                End If
            
            Next i
    
        ' Reset last row variable for next worksheet
        last_row = 0
    
    Next ws

End Sub


