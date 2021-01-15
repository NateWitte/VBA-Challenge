Attribute VB_Name = "Module1"
Sub SummarizeStock()
    'Starting Variables
    Dim currentsum As Double
    Dim numberofrows As Double
    Dim currentticker As String
    Dim startingvalue As Double
    Dim endvalue As Double
    Dim numbertickers As Double
    'BONUS Variables
    Dim currentMaxIncrease As Double
    Dim currentWorstDecrease As Double
    Dim currentMaxTotal As Double
    Dim MaxIncreaseTicker As String
    Dim WorstDecreaseTicker As String
    Dim MaxTotalTicker As String
    
    For Each ws In Worksheets
        currentsum = 0
        numberofrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        currentticker = ws.Cells(2, 1).Value
        startingvalue = ws.Cells(2, 3).Value
        'MsgBox (numberofrows)
        
        'Set up report table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        numbertickers = 2 'We start our table at row 2
        
        
        For i = 2 To numberofrows
            currentsum = currentsum + ws.Cells(i, 7).Value
            currentticker = ws.Cells(i, 1).Value
            nextticker = ws.Cells(i + 1, 1).Value
            If nextticker <> currentticker Then
                endvalue = ws.Cells(i, 6).Value
                yearlychange = endvalue - startingvalue
                If yearlychange > 0 Then
                    ws.Cells(numbertickers, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(numbertickers, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(numbertickers, 9).Value = currentticker
                ws.Cells(numbertickers, 10).Value = yearlychange
                If startingvalue = 0 Then
                    ws.Cells(numbertickers, 11).Value = "NA"
                Else
                    ws.Cells(numbertickers, 11).Value = yearlychange / startingvalue
                End If
                ws.Cells(numbertickers, 12).Value = currentsum
                'Reset everything and update
                currentsum = 0
                startingvalue = ws.Cells(i + 1, 3).Value
                numbertickers = numbertickers + 1
            End If
        Next i
        'Make Column 11 percentages
        ws.Range("K2:K" & numbertickers).NumberFormat = "0.00%"
        
        'BONUS Part
        currentMaxIncrease = 0
        currentWorstDecrease = 0
        currentMaxTotal = 0
        For i = 2 To numbertickers
            If ws.Cells(i, 11) <> "NA" Then
                If ws.Cells(i, 11).Value > currentMaxIncrease Then
                    currentMaxIncrease = ws.Cells(i, 11).Value
                    MaxIncreaseTicker = ws.Cells(i, 9).Value
                End If
                If ws.Cells(i, 11).Value < currentWorstDecrease Then
                    currentWorstDecrease = ws.Cells(i, 11).Value
                    WorstDecreaseTicker = ws.Cells(i, 9).Value
                End If
                If ws.Cells(i, 12).Value > currentMaxTotal Then
                    currentMaxTotal = ws.Cells(i, 12).Value
                    MaxTotalTicker = ws.Cells(i, 9).Value
                End If
            End If
        Next i
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(2, 17).Value = currentMaxIncrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = WorstDecreaseTicker
        ws.Cells(3, 17).Value = currentWorstDecrease
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxTotalTicker
        ws.Cells(4, 17).Value = currentMaxTotal
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Next ws
    
            
            
        
        
End Sub
