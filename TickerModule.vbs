Attribute VB_Name = "Module1"
Sub Ticker()

    'Declare variables
    Dim Ticker, Header As String
    Dim yearChange, percentChange, totalVolume, yearStart, yearEnd, tickerCount As Long
    Ticker = "" 'initialize Ticker
    tickerCount = 1
    'Set formatting
    Range("J:J").NumberFormat = "0.00"
    Range("K:K").NumberFormat = "0.00%"
    'Initialize Header
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    
    
    'For i = 2 To 2000 for loop for coding to make testing easier
    i = 2
    While Not (IsEmpty(Cells(i, 1))) 'Run while cell is not empty
        If (Cells(i, 1) <> Ticker) Then
            If (tickerCount > 1) Then
                yearEnd = Cells(i - 1, 5)
                yearChange = yearEnd - yearStart
                percentChange = yearChange / yearStart
                Cells(tickerCount, 9) = Ticker
                Cells(tickerCount, 10) = yearChange
                Cells(tickerCount, 11) = percentChange
                Cells(tickerCount, 12) = totalVolume
                
                If (yearChange > 0) Then
                    Cells(tickerCount, 10).Interior.ColorIndex = 4
                Else: Cells(tickerCount, 10).Interior.ColorIndex = 3
                End If
                
            End If
                
            tickerCount = tickerCount + 1
            Ticker = Cells(i, 1)
            yearStart = Cells(i, 5)
            totalVolume = Cells(i, 7)
            'MsgBox ("New ticker: " + Ticker)
        
        Else
            totalVolume = totalVolume + Cells(i, 7)
            
        End If
        i = i + 1
    
    Wend

End Sub

