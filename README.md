# VBA_Challenge

This is the vba code for the VBA Challenge for MSU Data Analytics Bootcamp.

This is code pretaining to a spreedsheet regarding stocks and summarizing the various information given for each stock.

One sub labeled "run code" will loop through the given information to add up the total volume of stocks, the value change for the year, and the percent change of value for the year. All that informaton is posted in a "summary table". Then another loop will run going through the "summary table" and further pick out the largest percent increase, largest percent decrease, and the largest volume of stocks and post those found stocks in another table reffered to as "summary of summary table". There is also some formatting applied. In the "summary table" on the "yearly change" and "percent change" columns it will color the cells green for a positive change and red for a negative change. It also autofits all column sizes at the end for better readability.

This "run code" is called in another sub called "SummarizeAllYears" which has the function of looping through all sheets. It will call for "run code" select the next sheet and then loop until all sheets are completed.


Running this will get the below result in the proper sheet. Access the spreadsheet here: https://drive.google.com/file/d/1UBLxWyRygqjO1BJcDKf3ukyRQk187ewZ/view?usp=sharing

[Click here to view the code!]: https://github.com/gmarciani](https://github.com/brown016340/VBA_Challenge/blob/main/Module2.bas

![image](https://github.com/brown016340/VBA_Challange/assets/142126077/fd57eac7-2c65-496f-94ca-9b1bb0bec15e)
