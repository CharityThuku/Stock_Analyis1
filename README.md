# Green Stock Analysis 

## Overview of the project
- Using the Visual Basic Apllication in excel, we helped Steve analze a data set to determine if a stock was worth investing in, based on the stock's total daily volume and yearly return. After completing the first analysis, we reused the macros created in the first analysis to to analyze 11 more stock to see which stock would be the smartest option.

## Purpose
- The purpose of this project is to learn how to use VBA to automate tasks and perform function that organize and analyze large datasets. By creating and running macros in VBA, we were able to analyze the large dataset into readable date that could depict how a stock performed in a certain year. In order to make the analysis run faster and efficiently, we learned how to refactor that code. 

## Results
- Refactoring means changing the structure of a program so that the functionality is not changed. It is done to make the code more efficient and maintanable. It is done by looping through the data at one time and collect all of the needed information. 
- Refactored Code

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

      'Get number of rows to loop over
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      '1a) Create a ticker Index
      Dim tickerIndex As Single
      tickerIndex = 0

      '1b) Create three output arrays
      Dim tickerVolumes(12) As Long
      Dim tickerStartingPrices(12) As Single
      Dim tickerEndingPrices(12) As Single
    
       '2a) Create a loop to initialize the tickerVolumes to zero
       For i = 0 To 10
       tickerVolumes(i) = 0
    
       Next i
    
        '2b) Loop over all the rows in the spreadsheet.
       For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
             
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
             tickerIndex = tickerIndex + 1
             
            End If
            
        'End If
    
       Next i
    
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
       For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        
      Cells(4 + i, 1).Value = tickers(tickerIndex)
      Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
      Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

      Next i

- The refactored code uses arrays to make the analysis more efficient as there are many variables of the same type in the code. These include tickerVolume, tickerStartingPrices and tickerEndingPrices assigned to each ticker symbol before running the data set. By including the tickers, the analysis is completed much faster than using the nested loops.

## Run-Time for Each Method and yearValue

### Original Code
2017


![1st 2017](https://user-images.githubusercontent.com/85714314/124399940-8d86ad80-dce4-11eb-9d1f-76be5db79147.png)

2018


![1st 2018](https://user-images.githubusercontent.com/85714314/124399959-adb66c80-dce4-11eb-9ebb-2fc31c13e2ed.png)

### Refactored Code
2017


![RF 2017](https://user-images.githubusercontent.com/85714314/124399971-c6268700-dce4-11eb-99ce-5f0a8afe2ba8.png)

2018


![RF 2018](https://user-images.githubusercontent.com/85714314/124399986-e35b5580-dce4-11eb-9db9-f6d101ccc69f.png)


### Conclusion
- Based on the run-times, the refactored code runs faster by approximately 0.6 seconds.


## Summary
### What are the advantages or disadvantages of refactoring code?
 - The advantage of refactoring is that it makes the analysis faster and more efficient while maintaining the functionality of the code. A disadvantage is that the process can be time consuming and includes a lot of risk. While it does maintain the original base code, a simple mishap may take alot of time to fix.

### How do these pros and cons apply to refactoring the original VBA script?
 - A pro is that the code will be better organized, it will run faster and more efficiently in VBA. A con is that refactoring is risky. If the code is wrong, it's hard to revert steps, thus programmers should always save a copy of the original code incase anything was to go wrong.
