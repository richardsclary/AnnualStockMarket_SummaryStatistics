Sub WallStreetStocks():

'--------------------    
'FIRST SUMMARY TABLE VARIABLES
'--------------------

Dim ticker As String 'Current Ticker
Dim numberTickers As Integer 'Tickers Per Worksheet
Dim lastRowState As Long 'Last Row Per Workshet
Dim openingPrice As Double 'Annual Opening Price
Dim closingPrice As Double 'Annual Closing Price
Dim yearlyChange As Double 'Annual Price Change
Dim percentChange As Double 'Annual Percent Change
Dim totalStockVolume As Double 'Total Stock Volume

'--------------------
'SECOND SUMMARY TABLE VARIABLES
'--------------------

Dim greatestPercentIncrease As Double 'Greatest Percent Increase Per Year
Dim greatestPercentIncreaseTicker As String 'Associates Ticker With Greatest % Increase Value
Dim greatestPercentDecrease As Double 'Greatest Percent Decrease Per Year
Dim greatestPercentDecreaseTicker As String 'Associates Ticker With Greatest % Decrease Value
Dim greatestStockVolume As Double 'Greatest Stock Volume Value Per Year
Dim greatestStockVolumeTicker As String 'Associates Ticker With Greatest Stock Volume Per Year

'--------------------
'ITERATE OVER EACH WORKSHEET IN WORKBOOK
'--------------------

For Each ws In Worksheets

    ws.Activate 'Make current worksheet active
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row 'Determine last row of current worksheet

    '--------------------
    ' ADD HEADERS TO COLUMNS FOR FIRST SUMMARY TABLE
    '--------------------

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    '--------------------
    ' SET VARIABLE INITIAL VALUES IN SUMMARY TABLE ONE
    '--------------------

    numberTickers = 0
    ticker = ""
    yearlyChange = 0
    openingPrice = 0
    percentChange = 0
    totalStockVolume = 0
    
    '--------------------
    ' PARSE DATA
    '--------------------

    For i = 2 To lastRowState 'For loop through list of tickers (skipping header row)

        ticker = Cells(i, 1).Value 'Retrieve value of current ticker symbol
        
        ' Retrieve value for annual opening price
        If openingPrice = 0 Then
            openingPrice = Cells(i, 3).Value
        End If
        
        ' Calculate/Sum total stock value of current ticker symbol
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
        ' Transitioning to a different ticker symbol and increment the number of tickers in the list
        If Cells(i + 1, 1).Value <> ticker Then
            numberTickers = numberTickers + 1
            Cells(numberTickers + 1, 9) = ticker
            
            ' Retrieve value for annual closing price for ticker
            closingPrice = Cells(i, 6)
            
            ' Calculate yearly change value
            yearlyChange = closingPrice - openingPrice
            
            ' Add yearly change value to the appropriate cell in each worksheet.
            Cells(numberTickers + 1, 10).Value = yearlyChange
            
            '--------------------
            'CONDITIONAL FORMATTING OF CELLS
            '--------------------

            If yearlyChange > 0 Then
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 4 'Fills cell green
            ElseIf yearlyChange < 0 Then
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 3 'Fills cell red
            Else
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 6 'Fills cell yellow
            End If
            
            
            ' Calculate percent change value for ticker.
            If openingPrice = 0 Then
                percentChange = 0
            Else
                percentChange = (yearlyChange / openingPrice)
            End If
            
            ' Format percent change value as a percent.
            Cells(numberTickers + 1, 11).Value = Format(percentChange, "Percent")
             
            ' Set opening price back to 0 when we get to a different ticker in the list.
            openingPrice = 0
            
            ' Add total stock volume value to the appropriate cell in each worksheet.
            Cells(numberTickers + 1, 12).Value = totalStockVolume
            
            ' Reset value when a new ticker is reached
            totalStockVolume = 0
        End If
        
    Next i
    
    '--------------------
    ' ADD HEADERS TO SECOND SUMMARY SECTION
    '--------------------
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row 'Determine last row of worksheet
    
    '--------------------
    ' SET VARIABLE INITIAL VALUES IN SECOND SUMMARY TABLE (FIRST ROW IN THE LIST)
    '--------------------

    greatestPercentIncrease = Cells(2, 11).Value
    greatestPercentIncreaseTicker = Cells(2, 9).Value
    greatestPercentDecrease = Cells(2, 11).Value
    greatestPercentDecreaseTicker = Cells(2, 9).Value
    greatestStockVolume = Cells(2, 12).Value
    greatestStockVolumeTicker = Cells(2, 9).Value
    
    '-------------------
    ' PARSE DATA
    '-------------------


    For i = 2 To lastRowState 'For loop through list of tickers (skipping header row)
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatestPercentIncrease Then
            greatestPercentIncrease = Cells(i, 11).Value
            greatestPercentIncreaseTicker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatestPercentDecrease Then
            greatestPercentDecrease = Cells(i, 11).Value
            greatestPercentDecreaseTicker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatestStockVolume Then
            greatestStockVolume = Cells(i, 12).Value
            greatestStockVolumeTicker = Cells(i, 9).Value
        End If
        
    Next i
    
    '--------------------
    ' ADD VARIABLE VALUES TO SECOND SUMMARY TABLE
    '--------------------

    Range("P2").Value = Format(greatestPercentIncreaseTicker, "Percent")
    Range("Q2").Value = Format(greatestPercentIncrease, "Percent")
    Range("P3").Value = Format(greatestPercentDecreaseTicker, "Percent")
    Range("Q3").Value = Format(greatestPercentDecrease, "Percent")
    Range("P4").Value = greatestStockVolumeTicker
    Range("Q4").Value = greatestStockVolume
    
Next ws


End Sub
