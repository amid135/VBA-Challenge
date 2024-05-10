Attribute VB_Name = "Module1"
Option Explicit
Sub ChallengeVBA()
    Dim ws As Worksheet
    Dim CurrentTicker As String
    Dim NextTicker As String
    Dim TotalVolume As LongLong
    Dim Summary_Table_Row As Long
    Dim InputRow As Long
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim ChangeFrac As Double
    Dim QuartChange As Long
    Dim PercentChangeMax As Double
    Dim GreatestVolume As LongLong
    Dim PercentChangeMin As Double
    Dim TickerIncrease As String
    Dim TickerDecrease As String
    Dim TickerVolume As String
    
    For Each ws In ThisWorkbook.Worksheets
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add the Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Add the Rows for Ticker and Value
    ws.Cells(2, 15).Value = "Greated % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ' Set an initial variabls
    TotalVolume = 0
    Summary_Table_Row = 2
    OpeningPrice = ws.Cells(2, 3).Value
    PercentChangeMax = ws.Cells(2, 11).Value
    PercentChangeMin = ws.Cells(2, 11).Value
    GreatestVolume = ws.Cells(2, 12).Value
    TickerIncrease = ws.Cells(2, 9).Value
    TickerDecrease = ws.Cells(2, 9).Value
    TickerVolume = ws.Cells(2, 9).Value
    
    ' Loop through all tickers and volume
    For InputRow = 2 To LastRow
        CurrentTicker = ws.Cells(InputRow, 1).Value
        NextTicker = ws.Cells(InputRow + 1, 1).Value
        TotalVolume = TotalVolume + ws.Cells(InputRow, 7).Value
        
        ' Check for last row of current stock
        If NextTicker <> CurrentTicker Then
            ' Input
            ClosingPrice = ws.Cells(InputRow, 6).Value
            
            ' Calculations
            QuarterlyChange = ClosingPrice - OpeningPrice
            ChangeFrac = QuarterlyChange / OpeningPrice
            
            ' Output
            ws.Range("I" & Summary_Table_Row).Value = CurrentTicker
            ws.Range("J" & Summary_Table_Row).Value = QuarterlyChange
            ws.Range("K" & Summary_Table_Row).Value = FormatPercent(ChangeFrac)
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
            
            ' Prepare for next Stock
            TotalVolume = 0
            Summary_Table_Row = Summary_Table_Row + 1
            OpeningPrice = ws.Cells(InputRow + 1, 3).Value
        End If

    Next InputRow
    
    For QuartChange = 2 To LastRow
        If ws.Cells(QuartChange, 10) > 0 Then
            ws.Cells(QuartChange, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(QuartChange, 10) < 0 Then
            ws.Cells(QuartChange, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(QuartChange, 10).Interior.ColorIndex = 2
        End If
    Next QuartChange
        

    For InputRow = 2 To LastRow
        If ws.Cells(InputRow, 11).Value > PercentChangeMax Then
            PercentChangeMax = ws.Cells(InputRow, 11).Value
            TickerIncrease = ws.Cells(InputRow, 9).Value
        End If
    Next InputRow
    
    ws.Cells(2, 16).Value = TickerIncrease
    ws.Cells(2, 17).Value = FormatPercent(PercentChangeMax)
    
    
    For InputRow = 2 To LastRow
        If ws.Cells(InputRow, 11).Value < PercentChangeMin Then
            PercentChangeMin = ws.Cells(InputRow, 11).Value
            TickerDecrease = ws.Cells(InputRow, 9).Value
        End If
    Next InputRow
    
    ws.Cells(3, 16).Value = TickerDecrease
    ws.Cells(3, 17).Value = FormatPercent(PercentChangeMin)
    
    
    For InputRow = 2 To LastRow
        If Cells(InputRow, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(InputRow, 12).Value
        TickerVolume = ws.Cells(InputRow, 9).Value
    End If
   Next InputRow
   
    ws.Cells(4, 16).Value = TickerVolume
    ws.Cells(4, 17).Value = GreatestVolume
Next ws
    MsgBox "Done"
End Sub
