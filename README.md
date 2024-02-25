# Stock Market Data Analysis with VBA

## Background

In this homework assignment, I embarked on a journey to leverage VBA scripting to analyze generated stock market data. The objective was to create a script that could loop through all the stocks for one year and output essential information such as yearly change, percentage change, and total stock volume.

## VBA Script Function/Goals

1. **Script Creation**: I developed a VBA script (found in the Module 2 Challenge VBA Script.txt file) that generated the following information for each stock:
   - Ticker symbol
   - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
   - Percentage change from the opening price to the closing price
   - Total stock volume of the stock

2. **Additional Functionality**: I enhanced my script to identify and return the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

3. **Script Adaptation**: I made adjustments to my VBA script to enable it to run on every worksheet (i.e., every year) at once.

## Implementation

- **Conditional Formatting**: I utilized conditional formatting to highlight positive change in green and negative change in red, ensuring easy visualization of data trends.

## Other Considerations

- **Dataset**: I used the provided "alphabetical_testing.xlsx" dataset for development. This smaller dataset allowed for quicker testing, with the expectation that my code should run in under 3 to 5 minutes.
- **Consistency**: I ensured that my script performed consistently across all sheets, maintaining the ease and efficiency that VBA brings to repetitive tasks.

## Code Citations

VBA Data Types:
(Found in lines 4-15 of StockChecker7 Sub)

https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

VBA Color Index:
(Found in lines 78 and 81 of StockChecker7 Sub)

https://www.automateexcel.com/excel-formatting/color-reference-for-color-index/

Format Cells in VBA to a Percentage with 2 Decimal Places:
(Found in lines 73, 116, and 120 of StockChecker7 Sub)

https://www.exceldemy.com/vba-format-percentage-2-decimal-places/
