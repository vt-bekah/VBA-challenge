# VBA-challenge
This repository contains challenge files for UT DAV Bootcamp Module 2 VBA Scripting

# File Notes
* VBA_Scripts\StocksummaryPerTab.vbs is the script that implements the requirements documented within the Instructions section with a summary on each tab. (makes sense for Multiple_year_stock_data workbook)
   * Note: This script assumes (per TA guidance) the the data is sorted alphadetically by ticker name then chronologicaly by date. This assumption allows for lighter processer usage during the analysis.
* The Results_Screenshots folder contains screenshots of each tab after executing the script on Multiple_year_stock_data.
   * MultiYearStock_20##Results are for each tab of the Multiple_year_stock_data workbook from using StocksummaryPerTab.vbs on the file. 

* The Starter_Code folder contains the excel files and screenshots provided in BCS/Canvas for completing the challenge.

# References
The following references were used to identify various functions used within the script:
 * Index: https://www.automateexcel.com/formulas/return-address-highest-value-in-range/
 * Max & Min: https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction#methods
 * Number Formatting: https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.numberformat
 * Autofit Formatting: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

# Instructions

Create a script that loops through all the stocks for one year and outputs the following information:
   * The ticker symbol
   * Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
   * The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
   * The total stock volume of the stock. The result should match the following image: moderate_solution
   * Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image: hard_solution
   * Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

# Other Considerations
   * Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.
   * Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.
