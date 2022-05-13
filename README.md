# Stock Analysis with Excel VBA

## Refactoring Microsoft Excel VBA script to improve runtime and efficiency of existing code.

VBA Challenge Excel file: [VBA Challenge - Stock Analysis](https://github.com/awalindeep/Stock-Analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project

### Background of Project

Steve Just graduated with Finance degree and his parents wanted to be his first clients. Steves parents are passionate about green energy and decided to invest their money in DAQO New Energy Corp. 

Steve offered to look into DAQO stocks for his parents and wanted to diversify their funds. He wanted to analyze handful of green service energy stock and have created an excel stock data file for 2018 and 2018.

Steve has asked to us to help with his data analysis. We prepared a workbook for him to run the analysis on the click of a button but is now looking to do an in-depth research for his parents. He wants to expand the dataset to entire stock market over last few years.  To achieve this we will have to refactor our existing code for him.


###Purpose of project

The purpose of this project is to edit or refactor a Microsoft Excel VBA script to run the code faster. By refactoring the original code we will be adding new functionality that will collect stock information in the year 2017 and 2018 quicker and efficiently. Although the original code worked well with current dataset but would not be efficient for larger set of data. To reduce the time of result the we will be refactoring the the original code in this challenge.

### The Project Data

We will be using same data that Steve initially presented to run our run time analysis. This data  included two charts with stock information on 12 different stocks from 2017 and 2018 . The stock information contains a ticker value, the date the stock was issued, the open, close and adjusted closing prices. It also includes the highest and lowest price, along with the volume of each stock. The goal here for us is to refactor the starter code provided and loop though retrieve the ticker, the total daily volume, and the return on each stock one time and collect all the information. 

## Results

### Refactoring the Code

In order to make the code more efficient, we had to switch the nesting order of for loops. To do this, we had to modify the starter code to following 

**Step 1a :** Create tickerIndex variable and set it equal to zero before iterating over all the rows.

      '1a) Create a ticker Index
    
     tickerIndex = 0
     
**Step 1b:** Create three output arrays:  `tickerVolumes`,  `tickerStartingPrices`, and  `tickerEndingPrices`

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
**Step 2a:** Create a  `for`  loop to initialize the  `tickerVolumes`  to zero.
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
**Step 2b:** Create a  `for`  loop that will loop over all the rows in the spreadsheet.
   
    '2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
**Step 3a:** Inside the  `for`  loop in Step 2b, write a script that increases the current  `tickerVolumes`  (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
   
    '3a) Increase volume for current ticker
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
**Step 3b:** Write an  `if-then`  statement to check if the current row is the first row with the selected  `tickerIndex`. If it is, then assign the current starting price to the  `tickerStartingPrices`  variable.  
     
     '3b) Check if the current row is the first row with the selected tickerIndex assign current starting price.
        
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
     End If
  **Step 3c:** Write an  `if-then`  statement to check if the current row is the last row with the selected  `tickerIndex`. If it is, then assign the current closing price to the  `tickerEndingPrices`  variable.  
  
     '3c) Check if the current row is the last row with the selected ticker and assign closing price
        
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     End If
**Step 3d:** Write a script that increases the  `tickerIndex`  if the next row’s ticker doesn’t match the previous row’s ticker.
    
    '3d Increase the tickerIndex if next row doesn't match the previous row.
            
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If

    Next i
**Step 4:** Use a  `for`  loop to loop through your arrays (`tickers`,  `tickerVolumes`,  `tickerStartingPrices`, and  `tickerEndingPrices`) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
   
    '4) Loop through arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
        
     Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1     


### Measure the Performance 

### Run time for All Stock Analysis

***2017*** 
  ![2017 All Stock Analysis Run time](https://github.com/awalindeep/Stock-Analysis/blob/main/Resources/2017_All_Stock_Analysis.png)

***2018***
![2017 All Stock Analysis Run time](https://github.com/awalindeep/Stock-Analysis/blob/main/Resources/2018_All_Stock_Analysis.png)

### Run time for Refactored All Stock Analysis

***2017***
![2017 Refactored all stocks Analysis Run time](https://github.com/awalindeep/Stock-Analysis/blob/main/Resources/2017_Refactored_All_Stock_Analysis.png)

***2018***
![2018 Refactored all stocks Analysis Run time](https://github.com/awalindeep/Stock-Analysis/blob/main/Resources/2018_Refactored_All_Stock_Analysis.png)


##Summary

### Pros and Cons of Refactoring Code

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

### The Advantages of Refactoring Stock Analysis

The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. Attached below are the screenshots that indicate the run time for our new analysis.

