# Challenge

## Refactor an existing to code to analyze stock ticker performances

### Background:
The existing code constructed during the VBA module led to building a code to analyze the performance of 12 green energy stocks based on their yearly return for the years 2017 and 2018.  The yearly return is the percentage increase and decrease in stock price from the beginning of the year to the end of the year for each of the stocks.  Stocks are represented by their tickers which is the main array in the code.  The code scanned through all the rows of data from the beginning of the rows for each ticker (which is part of an array), until the end of the row representing that particular ticker.  The volume of trading during all the dates of the data is cumulated to demonstrate the trade volume.  Another output array with starting price and ending price for each of the ticker is used to help calculate the performance of the respective stock.

### Objective

Refactor the code to be more efficient either in use of computing resources. This would help accomplish sorting through a larger dataset relatively quicker than the existing code. 
      
* The code initially starts by a popping up an input box which helps select the year of analysis of the stocks.
* Activate the output worksheet with the required headers representing the columns of computed data.  In this case, the "Ticker" of the stock is to be placed in the first column of the output sheet followed by "Total Daily Volume" and the "Return" in the subsequent columns.  This is accomplished by using the Cells(row number, column number).Value = "Name of the header" syntax.
* The main header of the output sheet in the Cell A1 is dynamic based on the input year of the stocks to be analyzed.  The syntax is Range("cell address").Value = "All Stocks (" + yearValue + ")".  the yearValue is populated based on the input year initially entered.

* The 12 stocks are represented by tickers and are initialized as an array.  Each of the tickers in the array have a unique index number.
*  The arrays Starting Price of a particular stock ticker, the Ending Price, the cumulative volume of trade carried out for that particular stock are all initiated.
*  Another dynamic variable called tickerStart is also initiated which is representing the start of the row where the new ticker is present.  The tickers are assumed to be sorted in alphabetical order and hence the no ticker is repeated after its initial pass of the rows that it represents.

The initial step is to calculate the number of rows in the entire dataset.  Column A rows where the tickers are present are computed and stored as RowCount.

In the data set, the start of sorting through the rows is set to row 2 in column A.

The loop is initiated with the ticker array starting with 0 to 11.  As the loop starts with 0 which is the first stock ticker, the volume of trading is reset to 0 as well.  The active worksheet of course being set to the year selected.
A subloop which runs from the start of the each starting row of the ticker represented as tickerStart as initiated above (which is 2 for the first ticker and changes accordingly) in each loop does the following:
* Computes the cumulative volume of trades for each ticker by incrementing the volume by the initial volume which is 0 with the respective value at the volume address of each row that the ticker is representing.
* Within this subloop, a conditional statement which stores the starting ticker price based on the fact that the ticker value one row above is not equal to the ticker and the ticker in the current row is equal to the ticker is used.
* An additional value representing the starting row of the next ticker is also stored during this conditional statement.  This helps to rapidly commence sorting through the next ticker at exactly this row location instead of from the beginning of the row number 2 as in the first loop.  This is the tickerStart variable.
* Similarly, the subloop has an ending price of the current ticker if the current row value is the same as the ticker and the next row value is not the current ticker.
 
These three output arrays are stored and presented in the output sheet by activating it and placing the values in the appropriate cell locations along with their formats.  The performance of the stock is computed by the ending price/starting price -1 in a percentage format.

The main loop is then repeated 11 more time with the next tickers in the array by resetting the volume traded to 0 again and computing the volume traded, starting price, new ticker starting row and the ending price.

### Limitations
* The code is versatile enough to where it can be used for analysis of other years if needed.  However, the code is limited to the 12 tickers in the array.
* Tickers are in alphabetical order and the code will need modifications if tickers are randomly placed in the dataset.
