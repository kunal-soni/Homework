Sub StockAnalysis()
    'Make Code run for all worksheets
    Dim WS As Worksheet
    For Each WS In Worksheets
        'Activate Worksheet
        WS.Activate
        
        'Initialize Labels for analysis
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Initialize Labels for Hard assignment
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        
        'Initialize Variables
        Dim Ticker As String
        Dim TotalStockVol As Double
        Dim TickerCounter As Integer
        
        'Initialize Stock Volume and Ticker counter
        TotalStockVol = 0
        TickerCounter = 2
        
        'More variables
        Dim YearlyChange As Double
        Dim YearInit As Double
        Dim YearEnd As Double
        Dim PercentChange As Double
        
        'Initialize Initial Price
        YearInit = Range("C2").Value
        
        'Hard Assignment Variables
        Dim GrtPerInc As Double
        Dim GrtPerDec As Double
        Dim GrtTotalVol As Double
        Dim GrtIncTicker As String
        Dim GrtDecTicker As String
        Dim GrtVolTicker As String
        
        'Initialize Hard Assignment Variables
        GrtPerInc = 0
        GrtPerDec = 0
        GrtTotalVol = 0
        GrtIncTicker = Range("A2").Value
        GrtDecTicker = Range("A2").Value
        GrtVolTicker = Range("A2").Value
        
        'Find Last Row
        Dim LastRow As Double
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Loop through the list
        For i = 2 To LastRow
            If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
                
                'Assign ticker value of old Stock ticker to variable
                Ticker = Cells(i, 1).Value
                
                'Add stock volume to variable
                TotalStockVol = TotalStockVol + Cells(i, 7).Value
            
                'Store Year End Price to closing value on last day
                YearEnd = Range("F" & i).Value
                YearlyChange = YearEnd - YearInit
            
                'Calculate Percent Change
                'This will be decimal value and style will be updated to percent
                If (YearInit = 0) Then
                  PercentChange = 0
                Else
                PercentChange = YearlyChange / YearInit
                End If
            
                'Store Ticker variable to new table
                Range("I" & TickerCounter).Value = Ticker
                
                'Store volume to table
                Range("L" & TickerCounter).Value = TotalStockVol
                
                'Store Price change to table and fill color accordingly
                Range("J" & TickerCounter).Value = YearlyChange
                If (YearlyChange > 0) Then
                    Range("J" & TickerCounter).Interior.ColorIndex = 4
                Else
                    Range("J" & TickerCounter).Interior.ColorIndex = 3
                End If
            
                'Store Percent change to table and Format
                Range("K" & TickerCounter).Value = PercentChange
                Range("K" & TickerCounter).NumberFormat = "0.00%"
                            
                'Increment Ticker counter
                TickerCounter = TickerCounter + 1
                            
                'Reset total stock volume to 0 and YearInit Price to next stock opening value
                TotalStockVol = 0
                YearInit = Range("C" & i + 1).Value
                        
                'Debugging help
                'MsgBox ("I=" & i)
            Else
                'Add Volume to total
                TotalStockVol = TotalStockVol + Cells(i, 7).Value
            End If
        Next i
        
        'Loop through all tickers in worksheet to determine greatest increase, decrease and volume
        For j = 2 To TickerCounter
            'Initialize current table variables
            Dim CurrentPerChange As Double
            Dim CurrentTicker As String
            Dim CurrentVol As Double
            
            'Assign values to current variables
            CurrentPerChange = Range("K" & j).Value
            CurrentTicker = Range("I" & j).Value
            CurrentVol = Range("L" & j).Value
            
            'Check conditions for the 3 greatest variables
            If (CurrentPerChange > GrtPerInc) Then
                GrtPerInc = CurrentPerChange
                GrtIncTicker = CurrentTicker
            End If
            If (CurrentPerChange < GrtPerDec) Then
                GrtPerDec = CurrentPerChange
                GrtDecTicker = CurrentTicker
            End If
            If (CurrentVol > GrtTotalVol) Then
                GrtTotalVol = CurrentVol
                GrtVolTicker = CurrentTicker
            End If
        Next j
        
        'Store these 3 greatest variables in new table
        Range("O2").Value = GrtIncTicker
        Range("P2").Value = GrtPerInc
        Range("P2").NumberFormat = "0.00%"
        Range("O3").Value = GrtDecTicker
        Range("P3").Value = GrtPerDec
        Range("P3").NumberFormat = "0.00%"
        Range("O4").Value = GrtVolTicker
        Range("P4").Value = GrtTotalVol
    Next WS
End Sub

