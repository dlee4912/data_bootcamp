Attribute VB_Name = "Module1"
Sub Button1_Click()
    Dim stockTicker As String
    Dim lastRow As Long
    Dim stockRow As Long
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    
    Dim greatestPercentInc As Double
    Dim greatestPercentDec As Double
    Dim greatestVolume As Double
    
    Dim gpiTicker As String
    Dim gpdTicker As String
    Dim gvTicker As String
    
    'Iterate through each worksheet
    
    For Each ws In Worksheets
    
        With ws
            ' Find the last row in a sheet and set the stock row for entry
            lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
            stockRow = 2
                
            ' Set the headers
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            ' Set the initial stock ticker symbole and opening price to the row
            stockTicker = .Range("A" & stockRow).Value
            openPrice = .Range("C" & stockRow).Value
            
            ' Go through all the rows for processing
            For r = 2 To lastRow
                ' Update the total trade volume
                totalVolume = totalVolume + .Range("G" & r).Value
                
                ' Check if the next row is still wtih the same ticker or not
                If (.Range("A" & r + 1).Value <> stockTicker) Then
                    
                    ' Grab the closing price
                    closePrice = .Range("F" & r).Value
                    
                    ' Add the entry to the summary table
                    .Range("I" & stockRow).Value = stockTicker
                    .Range("J" & stockRow).Value = closePrice - openPrice
                    .Range("K" & stockRow).Value = (closePrice - openPrice) / openPrice
                    .Range("L" & stockRow).Value = totalVolume
                    
                    ' Add formatting
                    If (.Range("J" & stockRow).Value > 0) Then
                        .Range("J" & stockRow).Interior.ColorIndex = 4
                    ElseIf (.Range("J" & stockRow).Value < 0) Then
                        .Range("J" & stockRow).Interior.ColorIndex = 3
                    End If
                    
                    .Range("K" & stockRow).NumberFormat = "0.00%"
                                        
                    ' Reset the ticker, opening price, and volume
                    stockTicker = .Range("A" & r + 1).Value
                    openPrice = .Range("C" & r + 1).Value
                    totalVolume = 0
                    
                    ' Increment the stock row
                    stockRow = stockRow + 1
                    
                End If
                
            Next r
            
        ' Format columns
        .Columns("I:L").AutoFit
        
        'Set the initial values to the first row in the summary table
        greatestPercentInc = .Range("K2").Value
        greatestPercentDec = .Range("K2").Value
        greatestVolume = .Range("L2").Value
        
        gpiTicker = .Range("I2").Value
        gpdTicker = .Range("I2").Value
        gvTicker = .Range("I2").Value
        
        ' Iterate through the summary table
        For r = 3 To stockRow
            ' Compare if the % is greater than the current value
            If (.Range("K" & r).Value > greatestPercentInc) Then
                greatestPercentInc = .Range("K" & r).Value
                gpiTicker = .Range("I" & r).Value
            End If
            
            ' Compare if the % is less than the current value
            If (.Range("K" & r).Value < greatestPercentDec) Then
                greatestPercentDec = .Range("K" & r).Value
                gpdTicker = .Range("I" & r).Value
            End If
            
            ' Compare if the volume is greater than the current value
            If (.Range("L" & r).Value > greatestVolume) Then
                greatestVolume = .Range("L" & r).Value
                gvTicker = .Range("I" & r).Value
            End If
            
        Next r
            
        ' Create the headers
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        
        ' Write out the values
        .Range("P2").Value = gpiTicker
        .Range("P3").Value = gpdTicker
        .Range("P4").Value = gvTicker
        
        .Range("Q2").Value = greatestPercentInc
        .Range("Q3").Value = greatestPercentDec
        .Range("Q4").Value = greatestVolume
        
        .Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Format columns
        .Columns("O:Q").AutoFit
        
        End With
    
    Next
End Sub
Sub clearSheets()
    For Each ws In Worksheets
    
        With ws
            .Range("I:Q").EntireColumn.ClearContents
            .Range("I:Q").EntireColumn.ClearFormats
        End With
    
    Next
End Sub
