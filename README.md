# challenge2VBA: The VBA of Wall Street


## Summarizing stock market data using VBA Scripts

For this Challenge, I followed the instructions provided to calculate each Stock Tickers performance over the course of each year that data was provided for (2018, 2019, 2020).

Not only was I able to utilize VBA to calculate the Yearly Change, Percent Change and Total Stock Volume over the course of each year, I also used the script to conditionally format the values for Yearly Change and Percent Change, so that a user can easily discern whether a stock's value increased or decreased over the given year.

Further analysis is offered in the form of a small table that provides the ticker which experienced each of the following:
1. Greatest % Increase
2. Greatest % Decrease
3. Greatest Total Volume

This further allows the user to quickly assess some of the outliers within each year's dataset.

One of my favourite aspects of the functionality of the script is it's ability to perfectly replicate the process across each worksheet in a workbook as long as the datasets are all formatted in the same way, with the following headings in the first row across the first seven columns of each sheet:

<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>
  

  


## Repo File Contents
- [SummaryModule.vbs](https://github.com/zmoloci/challenge2VBA/blob/main/SummaryModule.vbs)
  - Visual Basics Module containing script
- [alphabetical_testing-incomplete.xlsm](https://github.com/zmoloci/challenge2VBA/blob/main/alphabetical_testing-incomplete.xlsm)
  - Small Excel file with a portion of the original dataset. Used to test script and quickly troubleshoot issues
- [2018_Screenshot.png](https://github.com/zmoloci/challenge2VBA/blob/main/2018_Screenshot.png)
  - Screenshot of the result showing the 2018 worksheet
- [2019_Screenshot.png](https://github.com/zmoloci/challenge2VBA/blob/main/2019_Screenshot.png)
  - Screenshot of the result showing the 2019 worksheet
- [2020_Screenshot.png](https://github.com/zmoloci/challenge2VBA/blob/main/2020_Screenshot.png)
  - Screenshot of the result showing the 2020 worksheet
- [README.md](https://github.com/zmoloci/challenge2VBA/blob/main/README.md)
  - This file

Please feel free to add alternate modules to this repo, but do so in a new file. Please do not edit or remove the existing files.

==Contact me with any questions==
