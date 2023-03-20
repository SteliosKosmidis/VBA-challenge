Set up the necessary variables and data types to store the required information (ticker symbol, yearly change, percent change, total stock volume).

Loop through each row of data in the worksheet.

Check if the ticker symbol in the current row is the same as the previous row. If it is not, store the opening price and the ticker symbol.

Add the current row's stock volume to the total stock volume.

Check if the current row is the last row with data for the current ticker symbol. If it is, calculate the yearly change, percent change, and output all the information in the appropriate columns.

Use conditional formatting to highlight positive change in green and negative change in red.

Store the information for the greatest percent increase, greatest percent decrease, and greatest total volume as you loop through the rows.

Output this information at the end of the worksheet.

Repeat the above steps for every worksheet (that is, every year) in the workbook.