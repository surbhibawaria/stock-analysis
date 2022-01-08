# Stock Analysis

## Overview of the Project

Steve who is recently graduated with his finance degree, wanted to look into Daqo stocks for his parents, but he is concerned about diversifying their funds. He wants to analyze a handful of green energy stocks in addition to Daqo's stock and for that he has created an excel file containing the stock data he wants to analyze.

### Purpose and Analysis

The purpose of this analysis is to help Steve understand the market trends. Also, to do a little more research for his parents, as Steve wants to expand the dataset to include the entire stock market over the last few years.

Macros are created in VBA to analyze the stock data that would be easily accessible for Steve. _Total Daily Volume_ and yearly _Return_ is identified for each of the twelve stocks for the years 2017 and 2018. Conditional formatting is used to highlight the positive (through Green color) and negative (through Red color) returns. Interface is created so that at the click of a button, Steve can analyze the entire dataset. During the analysis two macros are created, the original script which was created from the module, and afterwards a refactored version of that script is created to determine whether refactoring the code made the VBA script run faster.

## Results

The stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script are compared.

### Stock Analysis Performance of 2017 vs 2018

For 2017, it can be seen that eleven out of the twelve stocks are positive and only one stock ticker 'TERP' is negative in terms of the yearly return. On the other hand for the year 2018, most of the stocks are negative and only two stock tickers, 'ENPH' and 'RUN' is positive. 'DQ', the initail stock that Steve's parents were interested in had a great year with 199.4% return in 2017. But in 2018, the yearly return dropped to 62.6%. 

To understand these stocks little better, let's take few examples of the stock tickers from these twelve stocks and comapre their valuation for both the years. 'AY' in 2017 have 8.9% return with a total daily volume of 136,070,900. In 2018, its return dropped to -7.3% with a total daily volume of 83,079,900 which is a decrease of 52,991,000 than in 2017. Similarly, 'CSIQ' in 2017, have return value of 33.1% with total daily volume of 310,592,800. Then in 2018, its return dropped in value by 16.3% with total daily volume of 200,879,900, which is about 110,000,000 less than it was in 2017. However on the other hand, 'VSLR' in 2017, have return value of 50.0% with total daily volume of 109,487,900. In 2018, its return drastically dropped in value by 3.5% with total daily volume of 136,539,100, which is about 27,051,200 more than it was in 2017 but still shows loss in return as we saw in the case of 'AY' and 'CSIQ'. This shows an uneven relationship between the total daily volume and return of the twelve stocks presented here and a major setback can be seen from year 2017 to 2018.

<img width="504" alt="AllStockAnalysis_2017" src="https://user-images.githubusercontent.com/95826875/148581869-e121ba67-fe36-4b7d-9836-78ee60ed709c.png">
<img width="504" alt="AllStockAnalysis_2018" src="https://user-images.githubusercontent.com/95826875/148581892-298adfb3-6ff5-40dd-a9c0-db57ee17fe51.png">

### Comparision between Execution Times of the Original Script and the Refactored Script

Refactoring is a vital part of any coding process. In this analysis, code refactoring played a major role. The original script ran in 0.2539062 seconds for the year 2017 and 0.2578125 seconds for the year 2018 as well. However, the refactored script ran in 0.0625 seconds for the year 2017 and 0.0703125 seconds for the year 2018. For both the years refactored code ran faster which can be seen below.

#### _Execution Time of the Original Script_

<img width="901" alt="RunTime_2017" src="https://user-images.githubusercontent.com/95826875/148605486-e4aca298-1ded-4e99-9279-ffc01763287b.png">
<img width="907" alt="RunTime_2018" src="https://user-images.githubusercontent.com/95826875/148605494-af6be53c-4395-4669-8cd1-9a59234e0f91.png">

#### _Execution Time of the Refactored Script_

<img width="767" alt="RunTimeRefactored_2017" src="https://user-images.githubusercontent.com/95826875/148605813-f8b4a7a3-1b32-4753-bf25-25f1b146590a.png">
<img width="723" alt="RunTimeRefactored_2018" src="https://user-images.githubusercontent.com/95826875/148605824-50a96cea-31cb-4573-a394-56abf92c57cc.png">

The illustration of original script vs refactored script will be discussed in the summary below.

## Summary

The process of restructuring code without changing or adding to its external behavior and functionality is known as _Code Refactoring_. The goal is to convert inefficient and overly complicated code into more efficient code that is preferably simpler and easier to understand.

### The Advantages and Disadvantages of Refactoring code

#### _Advantages_

1. Code refactoring makes the code more extensible, allowing for the addition of many more functions. It also aids in increasing the flexibility of the code, thereby increasing the capability of the code.

2. The code becomes cleaner, easier to understand or read, less complex, and easier to maintain after refactoring.

#### _Disadvantages_

1. One may have no idea how long it will take to complete the procedure of refactoring. It may also put one in a situation where they don't know where to go.

2. If something goes wrong while refactoring, one will have to spend a lot more time fixing it, and there's a good chance it'll go wrong due to the code's complexity.


### How does Pros and Cons apply to Refactoring the Original VBA Script

If we look at the codes below, the original script used nested loops while the refactored code used arrays. Although nested loops are an effective way of doing it, they can take longer to execute. Like here in the loop below, the code runs through each record of the data set once for each ticker category, on each stock's total volume, starting, and ending price. It looks certainly easier here as there are only twelve categories, which will not be the case every time. So, the nested loop is removed and arrays are used instead to reduce the time, and now the code access each data row only once.

#### _Original Script_

<img width="660" alt="Original_Script" src="https://user-images.githubusercontent.com/95826875/148611414-f909c683-91df-446a-9cd4-94b8cf00aad7.png">

#### _Refactored Script_

<img width="671" alt="Refactored_Script" src="https://user-images.githubusercontent.com/95826875/148611446-50fc0131-3ffa-40d2-8233-13d52f54242a.png">

As we saw the refactored script does run faster than the original script for both 2017 and 2018. Refactoring the original VBA script made the code more efficient, reduced the logarithmic complexity and made it look cleaner and simpler.

