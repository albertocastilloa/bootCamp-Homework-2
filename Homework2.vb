Sub getTickerAndVolume()
    Call performanceOff
    'Call function to get Ticker name and sum Volume
   Call allWorksheets
End Sub

'Call every page in Workbook
Function allWorksheets()
    Dim f_activeSheet As Worksheet
    For Each f_activeSheet In Worksheets
        'Change StatusBar
        Application.StatusBar = "Working on " + ActiveSheet.Name + " sheet..."
        f_activeSheet.Select
        Call getTickerVolume
    Next
    MsgBox("Script ended")
End Function

'Challenge 1: Easy
Function getTickerVolume()
    'Shut down temporaly excel function to improve script performance
    Call performanceOn
    'Retrieve names of Ticker 1 and 2
    Dim currentTicker1, currentTicker2 As String
    'Retrieve current and total Volume
    Dim currentVolume, totalVolume As Double
    'To control location where Ticker and Volume will crash
    Dim currentRow As Integer
    'Userful to find the open value of a Ticker
    Dim loopCounter As Integer
    'Last row in every WorkSheet
    Dim lastRow As Long
    'Initialize variables
    totalVolume = 0
    currentRow = 1
    loopCounter = -1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Loop to retrieve Tickers, compare values, sum volumes and place info
    'Total rows: 760192
    For i = 2 To lastRow
        currentTicker1 = Cells(i, 1).Value2
        currentTicker2 = Cells(i + 1, 1).Value2
        currentVolume = Cells(i, 7).Value2
        totalVolume = totalVolume + currentVolume
        loopCounter = loopCounter + 1
        
        If (currentTicker2 <> currentTicker1) Then
            currentRow = currentRow + 1
            Cells(currentRow, 9).Value = currentTicker1
            Cells(currentRow, 12).Value = totalVolume
            'Call getTicjerStockVolume - Challenge 2:Medium
            Call getTickerStockVolume(currentRow, i, loopCounter)
            'Reset variables
            totalVolume = 0
            loopCounter = -1
        End If
    Next i
    'Notify script is done
    Application.StatusBar = "Done"
    getGreatest (currentRow)
    'Active Excel funcion
    Call performanceOff
End Function

'Challenge 2: Medium
Function getTickerStockVolume(f_currentRow, f_i, f_loopCounter)
    'Hold both Open/Close price
    Dim openPrice, closePrice As Double
    'Retrieve Open and Close prices to get Yearly Change
    openPrice = Cells(f_i - f_loopCounter, 3).Value2
    closePrice = Cells(f_i, 6).Value2
    Cells(f_currentRow, 10).Value = closePrice - openPrice
    Cells(f_currentRow, 10).NumberFormat = "0.000000000"
    'Formating conditional according Yearly change value
    If (Cells(f_currentRow, 10).Value) >= 0 Then
        Cells(f_currentRow, 10).Interior.ColorIndex = 4
    Else
        Cells(f_currentRow, 10).Interior.ColorIndex = 3
    End If
    'Get Percent change
    If (openPrice = 0 Or closePrice = 0) Then
        Cells(f_currentRow, 11).Value = 0
    Else
        Cells(f_currentRow, 11).Value = closePrice / openPrice - 1
    End If
    Cells(f_currentRow, 11).NumberFormat = "0.00%"
End Function

'Challenge 3: Hard
Function getGreatest(countTotalRows)
    Dim tickerRows As Integer
    'Get resume of total values
    Dim currentPercent, maxPercent, currentStockVolume, maxStockVolume As Double
    'ticketRows is useful to build full range to find Max/Min values
    tickerRows = countTotalRows
    'Initialize variables
    maxPercent = 0
    minPercent = 0
    maxStockVolume = 0
    'Get Max/Min Percent Change and Total Stock Volume
    For i = 2 To tickerRows
        'Get Max Increase
        currentPercent = Cells(i, 11).Value2
        If currentPercent > maxPercent Then
            maxPercent = currentPercent
            Range("P2").Value = Cells(i, 9).Value2
            Range("Q2").Value = maxPercent
            Range("Q2").NumberFormat = "0.00%"
        End If
        'Get Min Decrease
        If currentPercent < minPercent Then
            minPercent = currentPercent
            Range("P3").Value = Cells(i, 9).Value2
            Range("Q3").Value = minPercent
            Range("Q3").NumberFormat = "0.00%"
        End If
        'Get Max Total Stock Volume
        currentStockVolume = Cells(i, 12).Value2
        If currentStockVolume > maxStockVolume Then
            maxStockVolume = currentStockVolume
            Range("P4").Value = Cells(i, 9).Value2
            Range("Q4").Value = maxStockVolume
        End If

    Next i
    'Funtion to build labels for new data
    Call buildColumnLabels
End Function

Function buildColumnLabels()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greates % Increase"
    Cells(3, 15).Value = "Greates % Decrease"
    Cells(4, 15).Value = "Greates Total Volume"
End Function

'A couple function to improve script performance
Function performanceOn()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Function

Function performanceOff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function