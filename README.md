# stock-analysis
stock market solar panel analysis with VBA

##Overview of Project
Steve wanted to analyze stocks from 2017 and 2018 to help his parents decide which stock to invest in. I have given Steve and Excel Macro-enabled workbook to analyze investment outcomes of 12 sotckes in 2017 and 2018. He liked the VBA script I produced due to the ability to get the information he needed with a click of a button. He is now interested in analyzing addtional stocks for more years of data. The objective of this project is to refactor the code I created to reduce the time it takes the computer to read through data to be prepared for his next assignemnt.

##Results
To refactor the code I created an array using "tickerIndex" to ask the computer to loop through data at a quicker pace than the original code. This change was used to find Total Volume, Starting Price and Ending Price as illustrated in following:

###Original Code to find Volume
If Cells(j, 1).Value = ticker Then
  totalVolume=totalVolume+Cells(j,8).Value
End if

###Refactored Code to find Volume
If Cell(i, 1).Value = ticker Then
  tickerVolumes(tickerIndex)=tickerVolumes(tickerIndex) + Cells(i, 8).Value
  
At first I through addtional lines of coding would take computer longer, but it made it quicker due to simplifying what I am asking the computer to do. The refactored code produced results over 90% faster than original. Final project [!All Stock Analysis with refactored and original codes]
Following show times of original code and refactored code for year year.

###2017 original code time stamp

###2017 refactored code time stamp

###2018 original code time stamp

###2018 refactored code time stamp

##Summary
The advantages of the refactored code is simply that it will run more data in a shorter amount of time. Also, the formatting is included in the macro and Steve will not have to run separate scripts to get a completed table. The disadvantage of the refactored script is that more work is put into the coding process and the script is not as clean and crisp to follow along and puts more necessity into the use of comments and white space. 

