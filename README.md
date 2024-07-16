# VBA-project
This VBA is designed to analyze quarterly stock data across multiple excel worksheets.
This subroutine iterates through each worksheet to summarize and calculate each quarterly stock based on the ticker symbol.
The data analysis loop
  Iterates through each row of data ('i') in the current worksheet.
  Checks if each row has a different ticker symbol.
  Identifies the ticker symbol and determines the range of rows for calculations.
The calculations
  Calculate the opening and closing price based on the first and last row of tickers data.
  Calculate quarterly and percent change and total volume for each ticker symbol.
The outputs
  are shown in the results columns labeled "tickers", "quarterly change", "percent change", and "total volume".
The 'Call to "Find Greatest Changes"' 
  finds the "GreatestChanges" subroutine to determine the greatest changes in percent and the output of the total volume.
The main purpose of this script is to analyze quarterly stock data across worksheets. This is useful for vizualizing and summarizing stock performances over time. 


