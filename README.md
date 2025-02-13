# VBA-challenge

A VBA script was created to analyze the provided stock data. For simplicity and ease of use, the VBA script was coded entirely on the 'This Workbook' VBA window, instead of writing separate codes on the individual sheets. Running the entirety of the consolidated code on this VBA window provides the desired result.

First, all the necessary variables were defined as Double, String or Long, as applicable. This can be viewed in the VBA script file. The For Each ws In ThisWorkbook.Sheets function was also used, to ensure the same code is applied to all the sheets in the workbook, without the need for any manual intervention. 

Next the column headings were defined for the desired output data. The first objective was to create a script to provide the ticker symbol, quarterly change, percentage change and total volume for the provided stock data.

To obtain all the individual stock tickers in the dataset a loop function was used. The End(xlUp).Row function was also used to assist with this, by identifying the last row in the dataset. The currentrow is set to 2, to ignore the headings in the first column. The Do While function was used to loop through each row of stock data and once a unique ticker is detected, Total volume for it is summed for all instances of that ticker. 

An If function is used to calculate the quarterly and percentage changes. A simple mathematical calculation is coded for both these metrics, and running the code calculates them automatically.

A simple If function is also used to calculate the greatest % increase/decrease and the maximum total volume. A separate If function is used for each requirement, and the code returns the ticker that satisfies the defined criteria. The cell ranges were also defined to output this data in columns N - P.