# Stock Summary Using VBA
## About the Project
This projects aim was to create a VBA script that could provide some basic analysis for Excel files that contain multiple sheets. The Excel files list companies and multiple entries for each company listing: 
- the date the data was taken
- the stock price at opening
- the highest price it reached that day
- the lowest price it reached that day
- the stock price at closing
- the volume of stocks
The VBA script creates a table adjacent to the Excel table. This new table provides an analysis of each company by listing:
- each companies yearly change in stock price and coloring them red if it has decreased in value and green if it increased in value
- a percentage of how much the stock has changed in price throughout the year
- the total traded stocks that year
## Using the VBA script
To utilize this script open one of the provided Excel sheets or another that uses the same structure, then go to the developer tab select the visual basic option. From within the visual basic menu either 'drag and drop' the YearlyStockCalc.vba into the side bar or select file and insert and add the YearlyStockCalc.vba file. After it has been added run the macro and wait for the message box that says "completed!".