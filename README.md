# stock-analysis
stock market solar panel analysis with VBA

## Overview of Project
Steve wanted to analyze stocks from 2017 and 2018 to help his parents decide which stock to invest in. I have given Steve an Excel Macro-enabled Workbook to analyze investment outcomes of 12 stocks in 2017 and 2018. He liked the VBA script I produced due to the ability to get the information he needed with a click of a button. He is now interested in analyzing addtional stocks for more years of data. The objective of this project is to refactor the code I created to reduce the time it takes the computer to read through data to be prepared for his next assignment.

## Results
To refactor the code I created an array using "tickerIndex" to ask the computer to loop through data at a quicker pace than the original code. This change was used to find Total Volume as shown below and used same coding to find starting and ending prices.

#### Example of original script to find volume
&ensp; If Cells(j, 1).Value = ticker Then
<br>&ensp; &ensp; totalVolume=totalVolume+Cells(j,8).Value
<br>&ensp; End if
#### Example of refactored script to find volume
&ensp; If Cell(i, 1).Value = ticker Then
<br>&ensp; &ensp; tickerVolumes(tickerIndex)=tickerVolumes(tickerIndex) + Cells(i, 8).Value
<br>&ensp; End if

At first, I thought addtional lines of coding would take computer longer, but it made it quicker due to simplifying what I am asking the computer to do. The refactored code produced results nearly 90% faster than original. Final script is coded to allow for any additional lines of data as well as addtional years given that worksheets of data are added to the Excel file for each year. [All Stock Analysis with refactored and original codes](https://github.com/LauraHaq/stock-analysis/blob/main/VBA_Challenge.xlsm) 
The following show screenshots of original code and refactored code for 2017 and then 2018.

### 2017 original code time stamp
![2017 Original](https://github.com/LauraHaq/stock-analysis/blob/main/original%202017%20time%20pic.png)
### 2017 refactored code time stamp
![2017 Refactored](https://github.com/LauraHaq/stock-analysis/blob/main/refactored%202017%20time%20pic.png)
### 2018 original code time stamp
![2018 Original](https://github.com/LauraHaq/stock-analysis/blob/main/original%202018%20time%20pic.png)
### 2018 refactored code time stamp
![2018 Refactored](https://github.com/LauraHaq/stock-analysis/blob/main/refactored%202018%20time%20pic.png)

## Summary
The advantages of the refactored code is simply that it will run more data in a shorter amount of time. Also, the formatting is included in the macro and Steve will not have to run separate scripts to get a completed table. The disadvantage of the refactored script is that more work is put into the coding process and the script is not as clean and crisp to follow along and puts more necessity into the use of comments and white space. Overall the project is successful and the script is ready to run more data for Steve.

