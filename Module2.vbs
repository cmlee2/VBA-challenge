Sub Module2()
Dim ws As Worksheet
'Loop through each ws
For Each ws In Worksheets

    'Defining Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Variable for Ticker Name
    Dim Ticker As String
    
    'Variable for Percent Difference
    Dim PercentDif As Double
    
    'Variable for Summary Table Row
    Dim SummaryRow As Integer
    SummaryRow = 2
    
    'Variable for Volume Total
    Dim VolumeTotal As LongLong
    VolumeTotal = 0
    
    'Variable for Close Price
    Dim CloseValue As Double
    
    'Variable for Highest Ticker and Highest Value Change
    Dim HighestTicker As String
    Dim HighestValue As Double
    HighestTicker = ws.Cells(2, 9)
    HighestValue = ws.Cells(2, 11)
    
    'Variable for Lowest Ticker and Lowest Value Change
    Dim LowestTicker As String
    Dim LowestValue As Double
    LowestTicker = ws.Cells(2, 9)
    LowestValue = ws.Cells(2, 11)
    
    'Variable for Volume Ticker and Largest Volume
    Dim VolumeTicker As String
    Dim VolumeValue As Double
    VolumeTicker = ws.Cells(2, 9)
    VolumeValue = ws.Cells(2, 12)
    
    'Tracking opening price
        Dim OpenValue As Double
        OpenValue = ws.Cells(2, 3).Value
    
    'Labeling columns for Ticker, Price, Percent, and Total Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Labeling Columns for Greatest % increase, Greatest % Decrease, and Greatest Total Volume
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Labeling Rows for Greatest % increase, Greatest % Decrease, and Greatest Total Volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
 
    'For loop through each row
    For i = 2 To LastRow
        
            'Conditional statement that will stop when the next tickers are not the same
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            
                'Setting the Ticker
                Ticker = ws.Cells(i, 1).Value
            
                'Track Closing Value
                CloseValue = ws.Cells(i, 6).Value
                
                'Calculate Difference in Open and Close over year
                Difference = CloseValue - OpenValue
                
                'Calculating Percent Difference
                PercentDif = (Difference / OpenValue)
                
                'Adding the current volume
                VolumeTotal = VolumeTotal + ws.Cells(i, 7)
            
               'Print Ticker in Summary Table
                ws.Range("I" & SummaryRow).Value = Ticker
                
                'Print Difference in Summary Total
                ws.Range("J" & SummaryRow).Value = Difference
                
                'Print Percentage in Summary Table
                ws.Range("K" & SummaryRow).Value = PercentDif
                
                
                'Print Total Volume in Summary Table
                ws.Range("L" & SummaryRow).Value = VolumeTotal
            
                'Add one to Summary Table row
                SummaryRow = SummaryRow + 1
                
                'Save new OpenValue
                OpenValue = ws.Cells(i + 1, 3).Value
                
                'Reset Total Volume
                VolumeTotal = 0
            
            'If next row is same ticker then...
            Else
            
                'Add to Total Volume
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
            
            End If
    Next i
    
    'New loop for summary table
    For i = 2 To (SummaryRow - 1)
        
        'Change style of summary rows to percentage
        ws.Cells(i, 11).NumberFormat = "0.00%"
         
        'Conditional that changes color if value is positive or negative
        If ws.Cells(i, 10) >= 0 Then
            ws.Cells(i, 10).Interior.Color = vbGreen
        Else
            ws.Cells(i, 10).Interior.Color = vbRed
        End If

        'Conditional that finds the greatest percent increase and ticker
        If ws.Cells(i + 1, 11) > HighestValue Then
            HighestTicker = ws.Cells(i + 1, 9)
            HighestValue = ws.Cells(i + 1, 11)
            ws.Cells(2, 17).Value = HighestValue
            ws.Cells(2, 16).Value = HighestTicker
        Else
        
        End If
        
        'Conditional that finds the greatest percent decrease and ticker
        If ws.Cells(i + 1, 11) < LowestValue Then
            LowestTicker = ws.Cells(i + 1, 9)
            LowestValue = ws.Cells(i + 1, 11)
            ws.Cells(3, 17).Value = LowestValue
            ws.Cells(3, 16).Value = LowestTicker
        Else
        
        End If
        
        'Conditional that finds the greatest volume value and ticker
        If ws.Cells(i + 1, 12) > VolumeValue Then
            VolumeTicker = ws.Cells(i + 1, 9)
            VolumeValue = ws.Cells(i + 1, 12)
            ws.Cells(4, 17).Value = VolumeValue
            ws.Cells(4, 16).Value = VolumeTicker
        Else
        
        End If
        
        
    Next i
    
'Formatting cells to show percentage variation
ws.Range("Q2:Q3").NumberFormat = "0.00%"

'Autofit cells for spacing
ws.Columns("A:Q").AutoFit

Next ws

End Sub





