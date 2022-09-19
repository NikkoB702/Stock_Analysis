# Stock_Analysis
Module 2 Stock Analysis
##### Purpose

The goal of this module was to learn how to use VBA and the reasons why we would choose to use VBA instead of working on the spreadsheet. The project cultimated in a refracting excercise. We extrated data from the years 2017 and 2018 to determine whether or not the list of twelve stocks are worth investing in.

##### The Data

The data originally presented included two charts with information on twelve different stock options. The stock information we received contained a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. We were then able to retrieve the ticker, the total daily volume, and the return on each stock.


## Results

##### Analysis

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


## Summary


##### Pros and Cons of Refactoring Code

Refactoring helps clean up code that is written and provides a more organized experience for anybody looking to read or write within the code. Refactoring can also help increase faster programming, debugging, and general software improvement. Some disadvantages may arise if the applications are too large or possibly not having the proper existing code. 

##### The Advantages of Refactoring Stock Analysis

When we use refactoring, we will see a decreased macro run time which can be benefit us especially if the code is searching through a large database. This will please anybody who needs to use this macro if they have time constraints.

![Screenshot 2022-09-18 172659](https://user-images.githubusercontent.com/80132877/190937154-8dd3bae6-710d-4786-ab57-8c6fbe08d261.png)

![Screenshot 2022-09-18 172815](https://user-images.githubusercontent.com/80132877/190937162-1545172a-7c4d-4086-ae28-5b26d47cd633.png)
