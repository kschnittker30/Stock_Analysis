# Project Overview
Create VBA macros to analyze stock data to calculate the total daily volume and yearly return. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Then refactor the VBA macros to make it more efficient by taking fewer steps and using less memory by looping through all the dataset at one time. The following task will be completed.
####    1. Refactor VBA code and measure performance.
####    2. A written analysis of results.
    
# Resources
Excel, VBA Macros

# Results

### Original VBA Macros:
![image](https://user-images.githubusercontent.com/99636479/167557328-59566aa1-0132-4f0b-898f-d314f758c186.png)
![image](https://user-images.githubusercontent.com/99636479/167557350-38ddc50f-eac1-44a5-81a0-42a2e8d9f4ac.png)



### Refactored VBA Macros:
![image](https://user-images.githubusercontent.com/99636479/167557392-476541a3-206c-4f5a-b25d-124fa48024fe.png)
![image](https://user-images.githubusercontent.com/99636479/167557411-64fdf79f-f191-48d4-9422-47f3ca154aa8.png)


# Summary
Refactoring a code can be advantageous by improving the logic of the code to run much faster and use less memory. Refactoring code can also keep your code clean and organized which makes it easier for other users to undestand.  The disadvantage of refactoring is that it can take longeer to develop and be more complicated.  More testing will be needed to ensure all the parts are working as expected. 

Refactoring the orginial VBA macro improved the runtime by around 84%. With only 12 stocks in the analysis, seconds between the two macros my seem minor by when the data is run over thousands of stocks several times in a period, the refactored code becomes increasing perferred. One disadvantage of refactoring the macros was the time and testing needed to run the macro smoothly.  
