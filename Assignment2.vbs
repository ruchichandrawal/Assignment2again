Sub LoopSheets()
    Dim i As Integer
    For i = 1 To 3
        Worksheets(i).Activate
        AnalyzeData
    Next i
End Sub

Sub AnalyzeData()
    'Dim i As Integer
    Dim lRow As Long
    Dim currentTicker As String
    Dim openAt As Single
    Dim closeAt As Single
    Dim volume As Double
    Dim printAtRow As Integer
    Dim yearlyChange As Single
    Dim percentChange As Single
    Dim formatRange As Range
    Dim minPercentChangeRow As Integer
    Dim maxPercentChangeRow As Integer
    Dim maxVolumeRow As Integer
    
    'Worksheets(1).Activate
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    currentTicker = Cells(2, 1).Value
    openAt = Cells(2, 3).Value
    closeAt = Cells(2, 6).Value
    volume = Cells(2, 7).Value
    
    
    'Result table
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    printAtRow = 2
    
    'Read and analyze stock data
    For i = 3 To lRow
        If Cells(i, 1).Value = currentTicker Then
            volume = volume + Cells(i, 7).Value
            closeAt = Cells(i, 6).Value
        Else
            'Ticker changed
            'Calculate previous stock result
            yearlyChange = closeAt - openAt
            percentChange = yearlyChange / openAt
            
            Cells(printAtRow, 9) = currentTicker
            Cells(printAtRow, 10) = yearlyChange
            Cells(printAtRow, 11) = percentChange
            Cells(printAtRow, 12) = volume
            printAtRow = printAtRow + 1
            
            'Reset variables
            currentTicker = Cells(i, 1).Value
            openAt = Cells(i, 3).Value
            closeAt = Cells(i, 6).Value
            volume = Cells(i, 7).Value
        End If
    Next i
    'print last result
    yearlyChange = closeAt - openAt
    percentChange = yearlyChange / openAt
    Cells(printAtRow, 9) = currentTicker
    Cells(printAtRow, 10) = yearlyChange
    Cells(printAtRow, 11) = percentChange
    Cells(printAtRow, 12) = volume
    
    'Format result
    Set formatRange = Range(Cells(2, 10), Cells(printAtRow, 11))
    Range("K:K").Style = "Percent"
    Range("K:K").NumberFormat = "0.0%"
    
    formatRange.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    formatRange.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    'Find and print min/max result
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    minPercentChangeRow = 2
    maxPercentChangeRow = 2
    maxVolumeRow = 2
    For i = 3 To printAtRow
        If Cells(i, 10).Value < Cells(minPercentChangeRow, 10) Then
            minPercentChangeRow = i
        End If
        If Cells(i, 10).Value > Cells(maxPercentChangeRow, 10) Then
            maxPercentChangeRow = i
        End If
        If Cells(i, 12).Value > Cells(maxVolumeRow, 12) Then
            maxVolumeRow = i
        End If
    Next i
    
    'Ticker
    Cells(2, 16) = Cells(maxPercentChangeRow, 9)
    Cells(3, 16) = Cells(minPercentChangeRow, 9)
    Cells(4, 16) = Cells(maxVolumeRow, 9)
    
    'Value
    Cells(2, 17) = Cells(maxPercentChangeRow, 11)
    Cells(3, 17) = Cells(minPercentChangeRow, 11)
    Cells(4, 17) = Cells(maxVolumeRow, 12)
    
    Range("Q2:Q3").Style = "Percent"
    Range("Q2:Q3").NumberFormat = "0.0%"
End Sub
