Option Explicit

Sub VBA_Challenge()

'cycle through every sheet of the workbook
Dim ws As Worksheet
For Each ws In Worksheets

    'Add summary table column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Add challange summary table rows/columns
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Declare summary variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim FirstOpenPrice As Double
    Dim LastClosePrice As Double
    
    'determin and set last row
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Track summary table row and set starting value
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    'cycle through every row of the data
    Dim i As Double
    For i = 2 To LastRow

        'Set first opening price if we didn't do so in a previous loop
        If FirstOpenPrice = 0 Then
            FirstOpenPrice = ws.Cells(i, 3).Value
        End If
        
        'if we have reached the final row of a stock...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'set stock ticker symbol
            Ticker = ws.Cells(i, 1).Value
            'set final closing price
            LastClosePrice = ws.Cells(i, 6).Value
            'calculate yearly change
            YearlyChange = LastClosePrice - FirstOpenPrice
            'Determine percent change for the year
            'in the case where there is no trade data, assign as 0
            'to avoid div/0 error
            If FirstOpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = YearlyChange / FirstOpenPrice
            End If
            'add to stock volume running total
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
              
            'Display values in summary table
            ws.Range("I" & SummaryTableRow).Value = Ticker
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
            ws.Range("K" & SummaryTableRow).Value = PercentChange
            ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
             
            'Conditionally format Yearly Change cell color
            If PercentChange < 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3 'red
            ElseIf PercentChange > 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4 'green
            End If
             
            'Reset variables and add 1 to summary table row
            Ticker = ""
            YearlyChange = 0
            PercentChange = 0
            TotalStockVolume = 0
            FirstOpenPrice = 0
            LastClosePrice = 0
            SummaryTableRow = SummaryTableRow + 1
        'if there are still more rows of this stock
        Else
            'add to stock volume running total
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
        End If

    Next i

    'challenge table code
    '=============================================================
    'declare table variable
    Dim MaxPercent As Double
    Dim MaxPercentTicker As String
    Dim MinPercent As Double
    Dim MinPercentTicker As String
    Dim MaxVolume As Double
    Dim MaxVolumeTicker As String
    
    MaxPercent = 0
    MinPercent = 0
    MaxVolume = 0

    'Loop through every row of the summary table
    Dim ChallengeRow As Integer  
    For ChallengeRow = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'check this row's percent change and compare it to the MIN so far
        If ws.Cells(ChallengeRow, 11).Value > MaxPercent Then
            'make this row's value the new max
            MaxPercent = ws.Cells(ChallengeRow, 11).Value
            'record the ticker of this row
            MaxPercentTicker = ws.Cells(ChallengeRow, 9).Value
        End If
        
        'check this row's percent change and compare it to the MIN so far
        If ws.Cells(ChallengeRow, 11).Value < MinPercent Then
            'make this row's value the new min
            MinPercent = ws.Cells(ChallengeRow, 11).Value
            'record the ticker of this row
            MinPercentTicker = ws.Cells(ChallengeRow, 9).Value
        End If
        
        'check this row's total volume and compare it to the MAX so far
        If ws.Cells(ChallengeRow, 12).Value > MaxVolume Then
            'make this row's value the new max
            MaxVolume = ws.Cells(ChallengeRow, 12).Value
            'record the ticker of this row
            MaxVolumeTicker = ws.Cells(ChallengeRow, 9).Value
        End If
    
    Next ChallengeRow
    
    'display results
    ws.Range("P2").Value = MaxPercentTicker
    ws.Range("Q2").Value = MaxPercent
    ws.Range("P3").Value = MinPercentTicker
    ws.Range("Q3").Value = MinPercent
    ws.Range("P4").Value = MaxVolumeTicker
    ws.Range("Q4").Value = MaxVolume
    '========================================================================
    
    'Resize coulmns to fit data
    ws.Range("I:Q").Columns.AutoFit
    'apply percentage format to percentage cells
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
Next ws

End Sub