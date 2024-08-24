Attribute VB_Name = "Module1"
Sub FindTickers()
    
    'Define the variables
    Dim ws As Worksheet
    Dim lastrow As Long, lastrowUnique As Long
    Dim openDate As Date, closeDate As Date
    Dim uniqueTickers As Collection
    Dim tickerArray() As Variant
    Dim tickerCount As Long
    Dim openValue As Double, closeValue As Double, quarterlyChange As Double
    Dim percentChange As Double, stockVolume As Double
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String
    Dim i As Long, j As Long, k As Long, l As Long
    Dim data As Variant, results() As Variant

    For Each ws In Worksheets
        
        'Determine the current sheet name
        Select Case ws.Name
            Case "Q1": openDate = DateSerial(2022, 1, 2): closeDate = DateSerial(2022, 3, 31)
            Case "Q2": openDate = DateSerial(2022, 4, 1): closeDate = DateSerial(2022, 6, 30)
            Case "Q3": openDate = DateSerial(2022, 7, 1): closeDate = DateSerial(2022, 9, 30)
            Case Else: openDate = DateSerial(2022, 10, 1): closeDate = DateSerial(2022, 12, 31)
        End Select

        'Find the last row of ticker data
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        'Load the data into an array
        data = ws.Range("A1:G" & lastrow).value

        'Loop through ticker data and put unique tickers into a collection
        Set uniqueTickers = New Collection
        On Error Resume Next
        For i = 2 To lastrow
            If Not IsEmpty(data(i, 1)) Then
                uniqueTickers.Add data(i, 1), CStr(data(i, 1))
            End If
        Next i
        On Error GoTo 0

        'Store the count of unique tickers
        tickerCount = uniqueTickers.Count
        
        'Resize ticker array and populate it with the unique tickers
        If tickerCount > 0 Then
            ReDim tickerArray(1 To tickerCount, 1 To 1)
            For i = 1 To tickerCount
                tickerArray(i, 1) = uniqueTickers(i)
            Next i
            
            'Put "Ticker" header and load them into column I
            ws.Cells(1, 9).value = "Ticker"
            ws.Range(ws.Cells(2, 9), ws.Cells(1 + tickerCount, 9)).value = tickerArray
            
            lastrowUnique = 1 + tickerCount
        End If
    

        'Initialize variables for max calculations
        maxIncrease = -1000
        maxDecrease = 1000
        maxVolume = 0

        'Prepare an array for output
        ReDim results(1 To lastrowUnique, 1 To 4)

        'Loop through unique tickers to calculate values
        For i = 2 To lastrowUnique
            openValue = 0
            closeValue = 0
            stockVolume = 0

            For j = 2 To lastrow
                If data(j, 1) = tickerArray(i - 1, 1) Then
                    'Open Date
                    If data(j, 2) = openDate Then
                        openValue = data(j, 3)
                    End If
                    'Close Date
                    If data(j, 2) = closeDate Then
                        closeValue = data(j, 6)
                    End If
                    'Total Stock Volume
                    stockVolume = stockVolume + data(j, 7)
                End If
            Next j

            'Calculate changes and update the max values
            If openValue <> 0 And closeValue <> 0 Then
                quarterlyChange = closeValue - openValue
                percentChange = (closeValue - openValue) / openValue
                results(i - 1, 1) = quarterlyChange
                results(i - 1, 2) = percentChange
                results(i - 1, 3) = stockVolume

                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = tickerArray(i - 1, 1)
                End If
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = tickerArray(i - 1, 1)
                End If
                If stockVolume > maxVolume Then
                    maxVolume = stockVolume
                    maxVolumeTicker = tickerArray(i - 1, 1)
                End If
            Else
                results(i - 1, 1) = "N/A"
                results(i - 1, 2) = "N/A"
                results(i - 1, 3) = stockVolume
            End If
        Next i

        'Output the results to the worksheet
        ws.Cells(1, 10).value = "Quarterly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Range("J2:J" & lastrowUnique).value = Application.Index(results, 0, 1)
        ws.Range("K2:K" & lastrowUnique).value = Application.Index(results, 0, 2)
        ws.Range("L2:L" & lastrowUnique).value = Application.Index(results, 0, 3)

        'Format the output cells
        ws.Range("J2:J" & lastrowUnique).NumberFormat = "0.00"
        ws.Range("K2:K" & lastrowUnique).NumberFormat = "0.00%"
        ws.Range("L2:L" & lastrowUnique).NumberFormat = "0"

        'Color code the quarterly changes
        For i = 2 To lastrowUnique
            If results(i - 1, 1) < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3 ' Red for negative
            ElseIf results(i - 1, 1) > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4 ' Green for positive
            Else
                ws.Cells(i, 10).Interior.ColorIndex = xlNone
            End If
        Next i

        'Output the greatest numbers
        ws.Range("O2").value = "Greatest % Increase"
        ws.Range("P1").value = "Ticker"
        ws.Range("P2").value = maxIncreaseTicker
        ws.Range("Q1").value = "Value"
        ws.Range("Q2").value = maxIncrease
        ws.Range("Q2").NumberFormat = "0.00%"

        ws.Range("O3").value = "Greatest % Decrease"
        ws.Range("P3").value = maxDecreaseTicker
        ws.Range("Q3").value = maxDecrease
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Range("O4").value = "Greatest Total Volume"
        ws.Range("P4").value = maxVolumeTicker
        ws.Range("Q4").value = maxVolume
    Next ws

End Sub
