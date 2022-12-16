# Stock-Analysis
## Overview of Project: 
After completing a workbook for Steve which helped him analyize several data sets from different stocks, he now wants to expand the dataset to include the entire stock market over the last few years to aid in the research for his parents.

## Explain the purpose of this analysis.

The purpose of this analysis is to "edit or refractor the existing code made for steve and then run the code to retrieve the same data. After the code has been run it will provide a timer which will help determine if the code now runs faster and more efficent.

## Analysis

The first step was to create an tickerIndex variable set to zero

Then I created 3 output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

  Dim tickerVolumes(12) As Long
    
  Dim tickerStartingPrices(12) As Single

  Dim tickerEndingPrices(12) As Single

Next, a for loop was created to initialize the tickerVolumes to the value of 0.

  For i = 0 To 11
    
     tickerVolumes(i) = 0

From there another for loop was created to loop over all the rows within the spreadsheet.

  For i = 2 To RowCount

Now a formula to increase the current tickerVolumes is created.

  tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

The next step needed to be an if-then statement to check if the current row is the first row with the selected tickerIndex. Then if it is, assign the current starting price to the tickerStartingPrices variable
 
  If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
       End If


Now an if-then statement is created to check if the current row is the last row with the selected tickerIndex. If it is, then the current closing price becomes the tickerEndingPrices variable.

  If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

  If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
Next, increase the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

   tickerIndex = tickerIndex + 1

Lastly, create a for loop to loop through the arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

## Results 
Found the time to be significantly faster than the previous code.

## compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Summary: 

### advantages of refactoring code?
Errors are found faster with a refractored code.
Refactoring code decreases micro time run.

### Disadvantages of refactoring code?
Time consuming whendealing with a larger code.
Able to makle mistakes more prevalent when removing data from the original code.

### How do these pros and cons apply to refactoring the original VBA script?
Dealing with a large amount of data and the different formulas needed can cause mistakes to easily hapen. With the bad comes the good as I found when refractoring this code. When I made a mistake after changing a large amount of code, the errors were found quickly and no change in data was made until the error was fixed.
