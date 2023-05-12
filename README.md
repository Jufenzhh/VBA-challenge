# VBA-challenge

VBA Stock Analysis 
This VBA macro analyses the data of stock through multiple years and creates a summary table that illustrates the ticker of the stock, its yearly change, percentage change, and its total stock volume. 

VBA Code Description 
The macro code starts by declaring the variables of the worksheet, ticker symbol, yearly change, percentage change, open and close prices. The code then loops through the worksheets and creates a new summary table that show all the ticker symbols, yearly change, percentage change, and total stock volume of the stock for each year, each worksheet is a new year.
Then the code shows if the stock for that year had  a positive yearly change or negative yearly change and assigns a color of green or red based on positive or negative. If there was a positive yearly change, yearly_change > 0, the cell in the new table would be green, and if it was negative, yearly change < 0, the cell would be red. 
Using the function Autofit from the columns of I through L which is where the new summary table was generated, the code could account for values to large for the cells, mostly in the case of total stock volume due to the size of numbers. 
From the new summary data of the stocks for a given year, the code generates the data of the greatest percentage increase and decrease, and greatest total volume. First through comparing the variables of percent_change, for greatest percentage increase and decrease, and for greatest total volume, the variable of total_volume. This creates a new summary table which shows the  greatest percentage increase and decrease, and  greatest total volume of stock for the given year. 
