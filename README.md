# Stock_Analysis
Module 2 Stock Analysis
#####Purpose

The goal of this module was to learn how to use VBA and the reasons why we would choose to use VBA instead of working on the spreadsheet. The project cultimated in a refracting excercise. We extrated data from the years 2017 and 2018 to determine whether or not the list of twelve stocks are worth investing in.

#####The Data

The data originally presented included two charts with information on twelve different stock options. The stock information we received contained a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. We were then able to retrieve the ticker, the total daily volume, and the return on each stock.


##Results

#####Analysis

I started with the VBA_Challenge template provided by the course instructors. I used Data Spell as the vessel to then copy and paste the code and instructions. I then proceeded with the refractoring process. Below is the code as described.

`Sub AllStocksAnalysisRefactored()


    Dim startTime As Single
    
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    

    'Initialize array of all tickers
    
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    '1a) Create a ticker Index
    
        tickerIndex = 0
        

    '1b) Create three output arrays
        
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single

        Dim tickerEndingPrices(12) As Single

    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.

        For i = 0 To 11
        
        tickerVolumes(i) = 0
        
        tickerStartingPrices(i) = 0
        
        tickerEndingPrices(i) = 0
        
        
Next i


        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
    
        '3a) Increase volume for current ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        
    End If
            
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        
     End If


            '3d Increase the tickerIndex.
            
                 If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            
            
        End If
        
    
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    
    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit

    dataRowStart = 4
    
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)



End Sub`
