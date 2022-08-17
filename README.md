# stock-analysis
Written Analysis of Results - Please reference document "Analysis of Deliverable" for summary, pictures and tables.

Purpose of Analysis:

The purpose of the analysis is for me to learn and get familiar with VBA coding.  In the exercise, the purpose is the provide a useful dataset
to help analyse the entire stock market over multiple years.  Finally, the purpose of the project is to frustrate me to the point I rather
hang a shower curtain.  

Results:

I started out with the code template provided by the module "challenge starter code".  I had a difficult time getting the VBA to run.  I had
alot of error messages "variable is already being used", "If end" not defined.  I switched gears and used the code from the homework in 
section 2.3.3 Reuse Code.  Early in the homework, I was able to get this code to work and run without any error messages.  I slowly added
the code from the challenge starter code to the "All Stocks Analysis" code to get results.  Its one paragraph, doesnt really explain the
amount of time it took me to get to that point.  Overall, it was a good strategy and it seemed to work., 

When comparing 2017 and 2018 tables, at a glance its easy to see which years and stocks performed better.  With the introduction of conditional
formatting (color formatting), its easy to see positive returns are green and negative returns are red.  From a glance at the color formatting
applied to the tables, 2017 was a much better year, returns where stronger.  The 2017 table is almost all green and the 2018 table is almost 
all red or negative.
 
 

Below are VBA codes I used to accomplish Deliverable 1 requirements.
1. The ticker index is set to equal to zero before looping over the rows:
For i = 0 To 11
    ticker = tickers(i)
    tickerIndex = tickers(i)
    totalVolume = 0
2.  Arrays are created for tickers, ticker volumes, ticker Starting Price and Ending Price:
Dim tickerVolumes As Long
Dim tickerStartingPrices As Single, tickerEndingPrices As Single
3. TickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices and Ending Prices arrays:
Worksheets("2018").Activate
'set initial volume to zero
 totalVolume = 0
For j = rowStart To rowEnd
If Cells (j, 1). Value = tickerIndex Then
4. The script loops through stock data:
If Cells(i, 1).Value = tickerIndex Then
 tickerVolumes = tickerVolumes + Cells(i, 8).Value
 End If
Worksheets("2018").Activate
    For j = 2 To RowEnd
    'Get the total volume for current ticker
    If Cells(j, 1).Value = ticker Then
   totalVolume = totalVolume + Cells(j, 8).Value
   End If
   'Get starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    startingPrice = Cells(j, 6).Value
    End If
    'Get ending price for current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    endingPrice = Cells(j, 6).Value
    End If
5. Code for formatting cells in the spreadsheet is working:
Table above, there are headers, bold is enabled, line style is present, headers are bold and numbers are formated with percentages.

6. There are comments to explain the purpose of the code, examples include:
'Increase volume for current ticker
 'Check if the current row is the first row with the selected tickerIndex.
  'check if the current row is the last row with the selected ticker
   'If the next row ticker doesnt match, increase the tickerIndex.
   'Increase the tickerIndex.
7. The output of the 2017 and 2018 tables match the module.  See referenced tables above.
8. The pop-up messages show elapsed run time:
The Run Time was relatively the same for both VBAs to run.  


 

Summary
The advantages of refactoring code in general are it allowed me to capture more data for the analysis.  The disadvantages are that to create the code was more complicated and time consuming.
The advantages and disadvantages of the original and refactored VBA script is that the original script was focused on one ticker DQ versus the challenge script that encompased all tickers.  To process the challenge script was time consuming, had more steps, overall was pretty tough.  The original script had less code to process and was less confusing.  Overall, I learned alot from the course and purchased a VBA book at Barnes and Nobels.  I actually like that coding and want to understand it more.  Half of me wishes we could just focus on VBA for the data analytics course.
