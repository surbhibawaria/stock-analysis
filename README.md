# Stock Analysis

## Overview of the Project

Steve who is recently graduated with his finance degree, wanted to look into Daqo stocks for his parents, but he is concerned about diversifying their funds. He wants to analyze a handful of green energy stocks in addition to Daqo's stock and for that he has created an excel file containing the stock data he wants to analyze.

### Purpose and Analysis

The purpose of this analysis is to help Steve understand the market trends. Also, to do a little more research for his parents, as Steve wants to expand the dataset to include the entire stock market over the last few years. And to refactor the code so that it doesn't take a long time to execute.

Macros are created in VBA to analyze the stock data that would be easily accessible for Steve. Total daily volume and yearly return is identified for each of the twelve stocks for the years 2017 and 2018. Conditional formatting is used to highlight the positive (Green) and negative (Red) returns. Interface is created so that at the click of a button, Steve can analyze an entire dataset. During the analysis two macros are created, the original script which was created from the module, and a refactored version of that script to determine whether refactoring the code made the VBA script run faster.

## Results

The stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script are compared.

### Stock analysis performance of 2017 vs 2018

For 2017, it can be seen that eleven out of the twelve stocks are positive and only one stock ticker 'TERP' is negative in terms of the yearly return. On the other hand for the year 2018, most of the stocks are negative and only two stock tickers, 'ENPH' and 'RUN' is positive. 'DQ', the initail stock that Steve's parents were interested in had a great year with 199.4% return in 2017. But in 2018, the yearly return dropped to 62.6%. 

To understand these stocks little better, let's take few examples of the stock tickers from these twelve stocks and comapre their valuation for both the years. 'AY' in 2017 have 8.9% return with a total daily volume of 136,070,900. In 2018, its return dropped to -7.3% with a total daily volume of 83,079,900 which is a decrease of 52,991,000 than in 2017. Similarly, 'CSIQ' in 2017, have return value of 33.1% with total daily volume of 310,592,800. Then in 2018, its return dropped in value by 16.3% with total daily volume of 200,879,900, which is about 110,000,000 less than it was in 2017. However on the other hand, 'VSLR' in 2017, have return value of 50.0% with total daily volume of 109,487,900. In 2018, its return drastically dropped in value by 3.5% with total daily volume of 136,539,100, which is about 27,051,200 more than it was in 2017 but still shows loss in return as we saw in the case of 'AY' and 'CSIQ'. This shows an uneven relationship between the total daily volume and return of the twelve stocks presented here and a major setback can be seen from year 2017 to 2018.

<img width="504" alt="AllStockAnalysis_2017" src="https://user-images.githubusercontent.com/95826875/148581869-e121ba67-fe36-4b7d-9836-78ee60ed709c.png">
<img width="504" alt="AllStockAnalysis_2018" src="https://user-images.githubusercontent.com/95826875/148581892-298adfb3-6ff5-40dd-a9c0-db57ee17fe51.png">

### Comparision between execution times of the original script and the refactored script

Refactoring is a vital part of any coding process. In this analysis code refactoring played a major role. The original code ran in 0.25 seconds for the year 2017 and 0.25 seconds for the year 2018 as well. However, the refactored code ran in 0.0625 seconds for the year 2017 and 0.0703125 seconds for the year 2018. 
## Summary

### The advantages or disadvantages of refactoring code

### How does pros and Cons apply to refactoring the original VBA script
