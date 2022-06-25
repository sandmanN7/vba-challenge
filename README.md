# vba-challenge
## Overview of Project

Using Visual Basic for Applications (VBA), I created a code that, at the click of a button, that could analyze stock data and output the information that was requested 
almost instantly. While this code is fine for a smaller number of stocks, it is not optimal for a much larger number of stocks. Thus, I set out to refactor the code
to make it more able to analyze more stocks and make it return information more quickly.

## Results

Initially, the refactoring process was difficult and I had to play with several variations of code, though I eventually found a functional solution.

Below is the refactored code and its run-times for 2017 and 2018 respectively:
```
'1a) Create a ticker Index
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long

Dim tickerStartingPrices(12) As Single

Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.

For i = 0 To 11
    tickerVolumes(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.
For j = 2 To RowCount

    '3a) Increase volume for current ticker

    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
     
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        
     End If

        '3d Increase the tickerIndex, loop prevents tickerIndex from going out of range
        
         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerIndex = tickerIndex + 1
            
         End If
        

Next j

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For k = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + k, 1).Value = tickers(k)
    
    Cells(4 + k, 2).Value = tickerVolumes(k)
    
    Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
    
Next k

```
![This is an image](https://github.com/sandmanN7/vba-challenge/blob/main/Resources/VBA_Challenge_2017.png)
![This is an image](https://github.com/sandmanN7/vba-challenge/blob/main/Resources/VBA_Challenge_2018.png)

Next is the original code, 2017 and 2018 runtimes: 
```
'set initial volume to zero
totalVolume = 0

Dim startingPrice As Double
Dim endingPrice As Double

'Establish the number of rows to loop over
rowStart = 2
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

'loop over all the rows
For i = rowStart To rowEnd

If Cells(i, 1).Value = "DQ" Then

'increase totalVolume by the value in the current row
totalVolume = totalVolume + Cells(i, 8).Value

End If

If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

startingPrice = Cells(i, 6).Value

End If

If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

endingPrice = Cells(i, 6).Value

End If

Next i

Worksheets("DQ Analysis").Activate
Cells(4, 1).Value = 2018
Cells(4, 2).Value = totalVolume
Cells(4, 3).Value = (endingPrice / startingPrice) - 1
```
![This is an image](https://github.com/sandmanN7/stock-analysis/blob/main/Green_stocks_2017.png)
![This is an image](https://github.com/sandmanN7/stock-analysis/blob/main/Green_stocks_2018.png)

Like the original code, the refactored code still uses conditonal loops to move through cells and output information. However, unlike the original
code, the refactored code uses arrays instead of just a single variable itself and has its outputs tied to directly to a tickerIndex in order to 
sort through information more quickly. As a result, the runtimes of the new code for both years was about one-sixth of the old code.

## Summary
### Pros and Cons of Refactoring in General
Refactoring allows for the creation of more efficient code. If successful, you are able to develop better code, have less bugs and are able to learn how to use a 
programming language more effectively. However, refactoring can cause a lot of issues. You inadvertently can create bugs and spend a more time than necessary trying 
to fix something than it is worth.

### Pros and Cons of the Old Code and the Refactored Code
The refactored code is a lot more time efficient, it takes a fraction of the time to complete than the old code and is better for going through larger amounts of data.
However, the refactored code is seemingly much more complex and it was much harder to get to work as running into bugs was much easier than one would think.
It took quite some time to make it work. On the otherside, while the old code was a lot slower, it was also much more simple and easier to get to work. 

