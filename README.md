# Stocks Analysis Refactoring

## Overview of Project

### Refactor Module 2 solution code to loop through all the data one time in order to collect the same information done in this module.  Then, determine whether refactoring the code succesfully made the VBA script run faster.

## Results

- Through the analysis of the output and performance we can get the following conclusios:

1.	Stock Performance: 
    - Although DQ has had a a stellar year in 2017 (+199.4%), it was the worst performer in 2018 (-62.6%).  So we can conclude this stock has a very high Beta to the stock markets, making it highly riscky.
    
    	- 2017
    	![2017](/VBA_Challenge_2017.png)
    	- 2018
    	![2018](/VBA_Challenge_2018.png)
    
2.	Execution Times:
	- There was a substantial improvement in the refcatured code compared to the original script, the code is runnig more than 4x faster. 
		
		- 2017
			- Before
			![Old2017](/No_refactoring_2017.png)
			- After
			![New2017](/VBA_Challenge_2017.png)
	
		- 2018
			- Before	
			![Old2018](/No_refactoring_2018.png)
			- After
			![New2018](/VBA_Challenge_2018.png)
			
- Code

Sub AllStocksAnalysisRefactored()

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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row's ticker doesn't match, increase the tickerIndex.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    
        'End If
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        'End If
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

End Sub




## Summary

- We can imply some advantages and disavantages of refactoring code:
	- Advantages:
		- Improve readbility;
		- Reduce complexity;
		- Improve source code maintanability;
		- Create a simpler, cleaner or more expressive internal architeture
		- Improve extensibility
		- Improve performance
	- Disavantages:
		- Time consuming
		
- These pros and cons applied to refactoring the original VBA script resumes in, although consuming some time in the process, it brought much more advantages in having a cleaner and more flexible code, which could be used in a much higher number of stocks, and improving the perdormance in more than 4 times.

