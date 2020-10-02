Attribute VB_Name = "Module11"
Sub InfoPull()
    
    Dim Ticker As String
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
    Dim CellVolume As LongLong
    Dim YearlyOpen As Variant
    Dim YearlyClose As Variant
    Dim YearlyChange As Variant

    Dim lastrow As Long
    Dim lastrow_summary As Long

    Dim SummaryMax As LongLong
    Dim SummaryIncrease As Long
    Dim SummaryDecrease As Long

    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    For Each ws In Worksheets
        
        ActiveSheet.Select
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'RENAME COLUMNS FOR SUMMARY TABLE AND ESTABLISH TABLE FILL LOOPS
        ws.Range("I1").Value = "Ticker Abv."
        Dim Ticker_row As Integer
        Ticker_row = 2
        
        ws.Range("J1").Value = "Yearly Change"
        Dim YearlyChange_row As Integer
        YearlyChange_row = 2
        
        ws.Range("K1").Value = "Percent Change"
        Dim PercentChange_row As Integer
        PercentChange_row = 2
        
        ws.Range("L1").Value = "Total Stock Volume"
        Dim TotalVolume_row As Integer
        TotalVolume_row = 2
        
        'RENAME COLUMS/ROWS FOR GREATEST VALUES TABLE
        'FOR CHALLENGE
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        '-----------------------------------------------------------'
        'TICKER PULL LOOP
        '-----------------------------------------------------------'

        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Ticker_row).Value = Ticker
                Ticker_row = Ticker_row + 1
                
            End If
            
        Next i
        
        YearlyOpen = ws.Cells(2, 3)
        
        'YEARLY OPEN/CLOSE/CHANGE/PERGENTAGE CHANGE LOOP
        
        For i = 2 To lastrow
                        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                YearlyClose = ws.Cells(i, 6).Value
            
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & YearlyChange_row).Value = YearlyChange
                YearlyChange_row = YearlyChange_row + 1
                
                If YearlyOpen <> 0 Then
                    PercentChange = (YearlyChange / YearlyOpen)
                    ws.Range("K" & PercentChange_row).Value = PercentChange
                    ws.Range("K" & PercentChange_row).NumberFormat = "0.00%"
                    PercentChange_row = PercentChange_row + 1
                End If
                
                If IsEmpty(ws.Cells(i + 1, 3)) = False Then
                    YearlyOpen = ws.Cells(i + 1, 3)
                End If
                
            End If
        
        Next i
        
        'TOTAL VOLUME CALCULATION AND LOOP
        
        TotalVolume = 0
    
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
                CellVolume = ws.Cells(i, 7).Value
                TotalVolume = TotalVolume + CellVolume
            
            End If
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ws.Range("L" & TotalVolume_row).Value = TotalVolume
                TotalVolume_row = TotalVolume_row + 1
                
                TotalVolume = 0
            
            End If
            
        Next i
        
        'CONDITIONAL FORMATTING
        
        lastrow_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary
        
            If ws.Cells(i, 10) < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            ElseIf ws.Cells(i, 10) >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
        
            End If
            
        Next i
        
        'FOR CHALLENGE -------------------------------------
        

        SummaryIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary))
        ws.Range("P2").Value = SummaryIncrease

        SummaryDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary))
        ws.Range("P3").Value = SummaryDecrease

        SummaryMax = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary))
        ws.Range("P4").Value = SummaryMax

        'ws.Range("O2").Value = Application.WorksheetFunction.Vlookup(SummaryIncrease, ws.Range("I2:L" & lastrow_summary), 1, False)
        'ws.Range("O2").Value = Application.WorksheetFunction.Vlookup(SummaryDecrease, ws.Range("I2:L" & lastrow_summary), 1, False)
        'ws.Range("O2").Value = Application.WorksheetFunction.Vlookup(SummaryMax, ws.Range("I2:L" & lastrow_summary), 1, False)
    Next ws
End Sub
