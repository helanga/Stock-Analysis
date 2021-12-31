# Stock-Analysis(Refacterd code) using VBA
## Overview of the Project
### Background of the project
### Purpose of the Project

## Results
### Performence Between 2017 & 2018(All Stock Analysis)
- Following is the output  tables results from Allstock analysis:

  ![](Resources/AllStockAnalysis2017Table.png)![](Resources/AllStockAnalysis2018Table.PNG)
  
 - Following is the performence between 2017 & 2018 worksheets according to All stock analysis script:

  ![](Resources/VBA_Challenge_2017.PNG)![](Resources/VBA_Challenge_2018.PNG)

 - We can say there is no significant difference in time to run between 2018 & 2017 worksheet because we wre using same code.

### Performence Between original script & refacterd script
- When refactering the script I added 3 output arrays to represent
  - tickerVolume
  - tickerStartingPrice
  - tickerEndingPrice
 - And use those arrays to calculate volume,starting price and ending price for each ticker.
 - Below image shows the part of the coding which I added mainly to  refactered code.

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
 
 
