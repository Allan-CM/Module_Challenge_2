# Excel Automation Showcase: Unleashing the Power of VBA

## About Project
VBA is a very powerful tool being used to increased efficency by automating a lot of repeptive tasks that you would manually do for a given dataset. 
After experimenting with VBA scripts, I've chosen to demonstrate the skills I've acquired by analyzing stock market data generated by the University of Toronto

## Built with
*	![Excel](https://img.shields.io/badge/Excel-63BE7B)

## Understanding the Project 
As mentioned previously this project analysis stock market data. For my project I decided to breakdown each step of my analysis into various script. 
- First I created a loop that runs through all the stock names which is commonly referred to as the ticker symbol to idenitfy all unique ticker symbols. Please see TickerSymbol Script
  - ![Ticker Symbol](https://github.com/Allan-CM/VBA-Excel-Showcase/blob/main/TicketScriptResult.png)
- Next I created a loop that runs through all stocks and outputs the difference between the closing and opening prioce for each stock at the end of the year to determine the yearly change
  - ![Yearly Change](https://github.com/Allan-CM/VBA-Excel-Showcase/blob/main/YearlyChangeScriptResult.png)
- Next, I created a loop that runs through all stock and outputs the percentage change for the year
  - ![Percentage Change](https://github.com/Allan-CM/VBA-Excel-Showcase/blob/main/PercentageChangeScriptResult.png)
- Next, I calculated the total stock volume for each stock 
  - ![TSV](https://github.com/Allan-CM/VBA-Excel-Showcase/blob/main/TotalStockVolumeScriptResult%20.png)
- Finally, the script runs through a loop to determine the ticker symbol and value for the categories of greatest increas/decrease and total volume 
  - ![Greatest Number](https://github.com/Allan-CM/VBA-Excel-Showcase/blob/main/GreatestNumbersScriptResult%20.png)

## Author contact
Allan Mathews - allancmathews@gmail.com

## Acknowledgements/References 
*To autofit cells contents code (ActiveSheet.UsedRange.EntireRow.AutoFit) was modeled from https://excelchamps.com/vba/autofit/)
*To format the percentage column and round to two decimal places code was modeled from (https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/)



