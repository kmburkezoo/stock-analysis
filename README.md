# VBA Stock Analysis
## Overview
This project analyzed the performance of several stocks over a particular year. Since the dataset included sheets for two years and the amount of information contained in each was extensive, VBA was used to prompt the user to input a year, then calculate volume and return for all included stocks within that year. The script also included conditional formatting to help differentiate positive from negatively performing stocks. The code was also refactored in order to analyze the efficiency of two different methods--the first, using a for loop to cycle through the 12 different tickers; the second, using an index variable and storing the values in a series of arrays before outputting them to the spreadsheet.

## Results
### Yearly Stock Performance
Analyzing the performance of all stocks together, rather than any single stock on its own, makes shared trends very clear: for most of these stocks, 2017 was a good year, and 2018 was a bad one. 

![2017 Stock Performance](https://github.com/kmburkezoo/stock-analysis/blob/main/Resources/2017_analysis.png) ![2018 Stock Performance](https://github.com/kmburkezoo/stock-analysis/blob/main/Resources/2018_analysis.png)

A drawback of the code as currently written is that it does not allow side-by-side analysis of the two years within the spreadsheet. However, viewing the two screenshots here, we can see that only two stocks had positive returns for both of the years under analysis. Based on the information presented here, there are three top candidates.
1. ENPH is the clear frontrunner, with exceptional performance in 2017 and good returns in 2018. 
2. RUN is also a strong contender. While it had only a 5% return in 2017, it fared comparably to ENPH in 2018.
3. SEDG is also worth watching. While its returns in 2018 were negative, they were only slightly so, and it did even better than ENPH in 2017.

### Execution Times
As mentioned in the overview, two different methods of analysis were compared:
1. Creating an array of tickers, then using a for loop to cycle through them and outputting the results to a spreadsheet before moving to the next ticker
```
    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

    'Loop through rows in the data.
        Worksheets("2018").Activate
        For j = 2 To RowCount

            'Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value

            End If

'            'Find the starting price for the current ticker.
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

                startingPrice = Cells(j, 6).Value
'
            End If
'            'Find the ending price for the current ticker.
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

                endingPrice = Cells(j, 6).Value

            End If
        Next j
'    'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
```
2. Creating an index variable to be used when looping through the array of tickers, then storing the results of each loop in an array before outputting all results to the worksheet.
```
'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    
    Dim tickerStartingPrices(11) As Single
    
    Dim tickerEndingPrices(11) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        
            '3a) Increase volume for current ticker
            'Find the total volume for the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            End If
                      
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
        
        Next i
        
'
        '3d Increase the tickerIndex.
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
```
Method #2 was faster, but not by a significant amount: method 1 took .92 seconds for 2017 and .98 for 2018, while method 2 took .63 seconds for 2017 and.65 seconds for 2018. In a program that takes minutes or hours to run, a 30% time savings can mean a great deal; however, since both versions took less than a second to run, the refactoring did not result in any real time savings.

## Summary
### Refactoring Code
### Refactoring _this_ Code
