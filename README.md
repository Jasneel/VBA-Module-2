# VBA-Module-2
Please see the VBA script . Thank You!-
Sub VBATickerLoop():

    
    
        Dim ws As Worksheet
        
        Dim i As Long
        Dim j As Long
        Dim TickerCounter As Long
        Dim LastTickerRowA As Long
        Dim LastTickerRowI As Long
        Dim PercentChange As Double
        Dim GreatestPercentIncrease As Double
        Dim GreatestPercentDecreaser As Double
        Dim GreatestTotalVolume As Double
        
        
        For Each ws In Worksheets ' to loop through worksheets
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        TickerCounter = 2
        j = 2
        
        
        LastTickerRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            
            For i = 2 To LastTickerRowA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                
                'Setting the cell color based on conditions
                    If ws.Cells(TickerCounter, 10).Value < 0 Then
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    End If
                    
                'Calculating Total Stock Volume
                ws.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
               
                TickerCounter = TickerCounter + 1
                j = i + 1
                
                End If
            
            Next i
            
        'Finding last non-blank cell in column I
        LastTickerRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        
        
        GreatestTotalVolume = ws.Cells(2, 12).Value
        GreatestPercentIncrease = ws.Cells(2, 11).Value
        GreatestPercentDecreaser = ws.Cells(2, 11).Value
        
            
            For i = 2 To LastTickerRowI
            
            
                If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestTotalVolume = GreatestTotalVolume
                
                End If
                
                
                If ws.Cells(i, 11).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestPercentIncrease = GreatestPercentIncrease
                
                End If
                
                '
                If ws.Cells(i, 11).Value < GreatestPercentDecreaser Then
                GreatestPercentDecreaser = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestPercentDecreaser = GreatestPercentDecreaser
                
                End If
                
            
            Next i
            
       
            
    Next ws
        
End Sub


