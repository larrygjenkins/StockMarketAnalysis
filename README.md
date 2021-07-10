# StockMarketAnalysis
![GitHub contributors](https://img.shields.io/github/contributors/larrygjenkins/larrygjenkins.github.io)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
## Description
The goal of this project was to create VBA scripts to analyze stock market data.

## Requirements
For each sheet in the Excel workbook, create a summary table that includes:

* The ticker symbol for each represented stock.
* The change from a stock's opening price for the year to its closing price for the year.
* The percent change for a stock's opening price for the year to its closing price for the year.
* The total stock volume for a given ticker symbol. 
* Conditional formatting for the Yearly Change column that highlighted positive changes in green and negative changes in red.

## Challenges
One of the stocks represented had "0" listed for all pricing data. This meant that calculations performed by the script could result in an Overflow Error because a formula would have included dividing by 0. 

To account for this, the script included a nested If statement to validate if the yearly open price was 0. If so, the percent change was automatically assigned a value of 0.

**Nested If Statement Validating Whether Year Open Price is 0**

        If yrOpenPrice = 0 Then
            percentChange = 0
                    
        Else
            percentChange = Round((((yrClosePrice - yrOpenPrice) / yrOpenPrice) * 100), 2)
                    
        End If

This solution served this particular data set, but it would have limitations if the opening price was 0 but the closing price was not. That scenario would require a decision on how to present a percentage change when the starting value is 0. (For example, would a starting price of 0 and an ending price of 1 be represented as a 100% change?)  

## Technologies Used
* Excel
* VBA

## Questions?
Contact me at the following locations:

* Email: <a href="mailto:larrygjenkins@gmail.com">larrygjenkins@gmail.com</a>
* GitHub: <a href="https://github.com/larrygjenkins">github.com/larrygjenkins</a>
* LinkedIn: <a href="https://www.linkedin.com/in/l-jenkins/">linkedin.com/in/l-jenkins</a>