# Stock-Analysis(Refacterd code) using VBA
## Overview of the Project
VBA is often used in the financial industry, so I worked with stock data in this module. 

### Background of the project
In this Project I'm helping Steve who just graduated with his finance degree.Steve's parents are passionate about green energy they believe fossil fuels will get used up there will be more and more dependability on alternative energy production.So his parents didn't do more reserch and decided to put all their money on "DAQO Energy Corp", DAQO's ticker symbol is DQ.

So for his parents he wants to analyze the data set for DQ's analysis.Steve wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded.The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve's parents are starting to pester him about Daqo's stock, and Steve wants to know how DQ performed in 2018. After Analyzing the DQ he realized that it  dropped over 63% in 2018,so then he started to analyze all the other stocks' Daily volume and return.

Finaly, I added buttons to analyze the entire dataset for different years with timer to calculate performance of the code.
Now Steve wanted a little more reserch for his parents, to expand the dataset and include the entire stock market over the last few years.



### Purpose of the Project

In this challenge, I refactored, the earlier script to produce more efficient code to analyze the entire stock market.

## Results
### Performence Between 2017 & 2018
- Following is the output  tables for Year 2017 & 2018:

  ![](Resources/AllStockAnalysis2017Table.png)![](Resources/AllStockAnalysis2018Table.PNG)
  
 - Below message box images shows the performence between 2017 & 2018 worksheets according to All stock analysis script:

  ![](Resources/VBA_Challenge_2017.PNG)![](Resources/VBA_Challenge_2018.PNG)

 - We can say there is no significant difference in time to run code years 2018 & 2017 worksheet because we wre using the same code.

### Performence Between original script & refacterd script
- When refactering the script I added 3 output arrays to represent:
  - tickerVolume
  - tickerStartingPrice
  - tickerEndingPrice
 - And use those arrays to calculate volume,starting price and ending price for each ticker.
 - Below image shows the part of the coding which I added mainly when refactoring the code.

 ![](Resources/refactcode.png)
 
 - Finally If we consider the time between refactered script & All Stock Analysis Stock it has significant different.
   - Time to run refacterd code for year 2017
   
      ![](Resources/VBA_Challenge_2017.PNG)![]()          
  
    - Time to run code before refactoring for year 2017
   
      ![](Resources/AllStockAnalysis2017time.png)
 
    - Time to run refacterd code for year 2018

      ![](Resources/VBA_Challenge_2018.PNG)
 
    - Time to run code before refactoring for year 2018
 
       ![](Resources/AllStockAnalysis2018time.PNG)
       
       
 ## Summery
 
 ### Advantages and disadvantages of refactoring code in general:
 #### Advantages
 Maintainability :After refactoring,the code is fresher,easier to understand or read,less complex and easier to maintain.
 #### Disadvantages
 Disadvantages of code refactoring :time consuming:You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go.
 
 ### Considering this project :
 If considering Refactored All Stock Analysis script it's performence increase significantly.So now Steve can use this workbook to analyse entire stock market and he can easily maintain it now with Arrays.
 Disadvantages,yes it's time consuming and we don't have exact idea what else we can do to improve performence and is it enogh the things which we did.
 
 
