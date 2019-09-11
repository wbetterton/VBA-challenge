Attribute VB_Name = "Module1"
Sub Ticker2()
    
        Dim TickerName As String
        
        Dim OpenAmount As Double
        
        Dim CloseAmount As Double
        
        Dim AmountChange As Double
        
        Dim Volume As Double
        
        Dim Percent As Double
               
        Dim Summary_Row As Integer
        Summary_Row = 2
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        OpenAmount = Cells(2, 3).Value
        
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                TickerName = Cells(i, 1).Value
                
                CloseAmount = Cells(i, 6).Value
                
                AmountChange = CloseAmount - OpenAmount
                
                OpenAmount = Cells(i + 1, 3).Value
                
                Percent = AmountChange / ((OpenAmount + CloseAmount) / 2)
                
                Volume = Volume + Cells(i, 7).Value
                
                Range("I" & Summary_Row).Value = TickerName
                
                Range("J" & Summary_Row).Value = AmountChange
                
                Range("K" & Summary_Row).Value = Percent
                    
                Range("K" & Summary_Row).NumberFormat = "0.00%"
                    
                Range("L" & Summary_Row).Value = Volume
                
                Summary_Row = Summary_Row + 1
                
                Volume = 0
                
            Else
                
                Volume = Volume + Cells(i, 7).Value
            
            End If
            
            If Cells(i, 11).Value < 0 Then
                Cells(i, 11).Interior.ColorIndex = 3
            Else
                Cells(i, 11).Interior.ColorIndex = 4
            End If
            
        Next i
   
End Sub

