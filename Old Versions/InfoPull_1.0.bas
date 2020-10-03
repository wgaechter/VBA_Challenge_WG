Attribute VB_Name = "Module1"
Sub InfoPull()
    
    Dim Ticker As String
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
    Dim CellVolume As LongLong
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim TradeDate As Date
    Dim lastrow As Long
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'RENAME COLUMNS FOR SUMMARY TABLE AND ESTABLISH TABLE FILL LOOPS
    Range("I1").Value = "Ticker Abv."
    Dim Ticker_row As Integer
    Ticker_row = 2
    
    Range("J1").Value = "Yearly Change"
    Dim YearlyChange_row As Integer
    YearlyChange_row = 2
    
    Range("K1").Value = "Percent Change"
    Dim PercentChange_row As Integer
    PercentChange_row = 2
    
    Range("L1").Value = "Total Stock Volume"
    Dim TotalVolume_row As Integer
    TotalVolume_row = 2
    
    '-----------------------------------------------------------'
    'TICKER PULL LOOP
    '-----------------------------------------------------------'
     
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Ticker = Cells(i, 1).Value
            Range("I" & Ticker_row).Value = Ticker
            Ticker_row = Ticker_row + 1
            
        End If
        
    Next i
    
    YearlyOpen = Cells(2, 3)
    
    'YEARLY OPEN/CLOSE/CHANGE/PERGENTAGE CHANGE LOOP
    
    For i = 2 To lastrow
                    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            YearlyClose = Cells(i, 6).Value
        
            YearlyChange = YearlyClose - YearlyOpen
            Range("J" & YearlyChange_row).Value = YearlyChange
            YearlyChange_row = YearlyChange_row + 1
            
            YearlyOpen = Cells(i + 1, 3)
            
            'PercentChange = (YearlyChange / YearlyOpen)
            'Range("K" & PercentChange_row).Value = PercentChange
            'Range("K" & PercentChange_row).NumberFormat = "0.00%"
            'PercentChange_row = PercentChange_row + 1
            
        End If
    
    Next i
    
    'TOTAL VOLUME CALCULATION AND LOOP
    
    TotalVolume = 0

    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        
            CellVolume = Cells(i, 7).Value
            TotalVolume = TotalVolume + CellVolume
        
        End If

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Range("L" & TotalVolume_row).Value = TotalVolume
            TotalVolume_row = TotalVolume_row + 1
            
            TotalVolume = 0
        
        End If
        
    Next i
    
    'CONDITIONAL FORMATTING
    
    For i = 2 To lastrow
    
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        
        ElseIf Cells(i, 10) >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
    
        End If
        
    Next i

End Sub
