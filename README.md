# stocks-analysis
Module 2 - VBA

## Project Overview

Refactoring is a key part of the coding process. Refactoring code is not about adding new functionality, you can make the code more efficient by taking fewer steps, using less memory, or improving the logic to make it easier for future users to read.

This project consists in two main milestones: 
- Refactor the solution from Module 2 to loop through all the data one time in order to collect the same information
- Determine whether refactoring the code successfully made the VBA script run faster

[green_stocks.xlsm file with Module 2 code](https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/green_stocks.xlsm) 

[VBA_Challenge.xlsm file with refactored code](https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/VBA_Challenge.xlsm)

#### Deliverable 1
- [x] 1. The tickerIndex is set equal to zero before looping over the rows
- [x] 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
- [x] 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays
- [x] 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
- [x] 5. Code for formatting the cells in the spreadsheet is working
- [x] 6. There are comments to explain the purpose of the code
- [x] 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module
- [x] 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

[VBA_Challenge_2017.png](https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/VBA_Challenge_2017.png) 

[VBA_Challenge_2018.png](https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/VBA_Challenge_2018.png)

#### Deliverable 2
- [x] 1. There is a title, and there are multiple paragraphs
- [x] 2. Each paragraph has a heading
- [x] 3. There are subheadings to break up text
- [x] 4. Links are working, and images are formatted and displayed where appropriate
- [x] 5. The purpose and background are well defined
- [x] 6. The analysis is well described with screenshots and code
- [x] 7. There is a detailed statement on the advantages and disadvantages of refactoring code in general
- [x] 8. There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script

## Resources 

- Excel - Version 16.54


## Results

While refactoring the code from Module 2 to achieve deliverable 1, several steps were followed to improve performance, more specifically in the **for** loops portion used to retrieve total volume, start and ending prices per ticker **while** populating the summary.

The original code uses two nested loops:

<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/01a111b9ca6f424ff7eda19ca26341efefff7f95/Resources/AllStockAnalysis.png">
</kbd>  &nbsp;
</p>

The refactored code separates the loops, the first retrieves the information:

<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/01a111b9ca6f424ff7eda19ca26341efefff7f95/Resources/AllStocksAnalysisRefactored1.png">
</kbd>  &nbsp;
</p>

And the second populates the summary per ticker:

<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/01a111b9ca6f424ff7eda19ca26341efefff7f95/Resources/AllStocksAnalysisRefactored2.png">
</kbd>  &nbsp;
</p>

The outcomes from both VBA scripts are the same: 

#### Outcome for 2017 (refactored code) 
<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/VBA_Challenge_2017.png">
</kbd>  &nbsp;
</p>


#### Outcome for 2018 (refactored code)
<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/VBA_Challenge_2018.png">
</kbd>  &nbsp;
</p>


The original code uses two nested loops and looks concise compared to the refactored one using arrays, but just beacause there's less steps and reading to do doesn't mean it's more efficient, the nested loops created an unnecessary complexity to the code that we can observe here:  




#### Two nested loops  
<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/nested_time_2017.png">
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/nested_time_2018.png">
</kbd>  &nbsp;
</p>

#### Refactored code
<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/time_2017.png">
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/time_2018.png">
</kbd>  &nbsp;
</p>

The runtime for the refactored code is around 4.5 times faster than the previous one done during Module 2. 

## Summary

Deliverable 1 made evident some of the advantages of refactoring code

Advantages of refactoring code:
How do these pros and cons apply to refactoring the original VBA script?
Helps Programming Faster
Improves the Design of Software
Refactoring Makes Software Easier to Understand
Refactoring Helps Finding Bugs
Helps Programming Faster


Disadvantages of refactoring code
- Requires money
- Requires time
- Can be risky when application is big
- Can be risky when developers do not understand much about the code


It's like picking the time to change an outdated cell phone, we usually consider if the benefits are going to be worth the price for the upgrade. Everything depends on the necessity and resources and impact of the improvements 


<kbd>
  <img src="https://github.com/GabrielaTuma/stocks-analysis/blob/a61ec8c980234f2bd876730c9ac8dd4a66a680be/Resources/year_question.png">
</kbd>  &nbsp;
</p>

