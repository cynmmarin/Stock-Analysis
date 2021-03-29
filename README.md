# Stocks Analysis

## Overview of Project
In this stocks analysis we will be helping Steve evaluate how profitable the DAOO New Energy Corp(DQ)stock has performed in the year 2017 and 2018. This will allow us to make an assessment of how lucrative this stock is in comparison to other green energy stocks. Through the application of code to automate analysis, we will be using VBA to make calculations that will help us analyze the ‘Total Daily Volume’ and ‘Return’ of the stocks for the given years. This will determine if investing in DQ is the optimal choice.

Our client Steve, has been hired by his parents to help them with investing their money in alternative energy production. Their focus is on green energy and the stock they have chosen to invest all their funds is DQ. Unfortunately, they have limited knowledge about the industry or the performance of the DQ stock in the market. For this reason, Steve has gathered the data and created the ***Green_Stocks*** Excel spreadsheet for 11 additional stocks for the year 2017 and 2018. The data includes the following information of all 12 stocks: ‘Opening’ price, ‘Closing’ price, ‘Highest’ traded price, ‘Lowest’ traded price, ‘Adjusted Closing’ price and the ‘Volume’ of shares that were traded on that given ‘Date’. 

## Analysis
Throught the application of codes on VBA, we will be able to use Steve’s data to analyze the traded ‘Total Daily Volume’ and ‘Return’ of all 12 stock ‘Tickers’ for the year 2017 and 2018. Thus determine the success or failure of the DQ stock in comparison to the other 11 stocks, and make an analytical assessment of the stock’s performance in the green energy industry.      

### All Stocks Analysis

The ***Green_Stocks*** spreadsheet has the raw data given to us by Steve. We will begin by running a ‘All Stocks Analysis’. Let’s start by opening VBA using the Developer tab. Once VBA is opened, we can go ahead and start coding. The first thing we want to do is start our subroutine ‘AllStocksAnalysis()’ this will help us keep track of our work. In order to make sure our calculations go to the correct worksheet, we will need to create a worksheet on Excel and label it ‘All Stocks Analysis’ once created, we can tell VBA to activate the worksheet and use this to output our data. 

Remember we want for Steve’s interaction with our analysis to be smooth as possible. For this we need to create a worksheet ‘yearValueAnalysis’ to code for a message (MsgBox), that will ask the user to let the program know what year they want the analysis to run. 

We want our analysis to use the data of the stocks’ performance in the years 2017 and 2018. Let’s go back to our ‘AllStockAnalysis’ worksheet. Let’s extract this information from the sheet and code a range that focus on all the different stocks available to us in this dataset for both given years. This code will allow us to output the results of both years. 

Our sheet is blank, and if we start coding without telling it where the data should be place, we will fail at giving Steve and organized summary of our findings. So, let’s create header rows. Let’s create a header for ‘Ticker’, ‘Total Daily Volume’ and ‘Return’, and define the cell we want this information to be located. This is where we want to output our result, now that we have an outline of what we want to find we can start using our data.

### Defending Variables and Creating Arrays

Next, we want to define our variables and assigning them a data type on VBA. We will do so by writing out the dimension ‘Dim’ the variable and then the type. In our case, the variables are ‘tickes’ which the type is a string, ‘startingPrice’ and ‘endingPrice’ which both are single. For the ‘tickers’, we want to create an array, a list of the 12 different stocks we want to analyze. 

### Creating Loops

Let’s create some loops, which are our way of telling the program to repeat lines of code, for certain number of times. Creating the loops will guarantee that the code we write will be applied to the data we need. For this we need to start by activating the ‘yearValue’ worksheet first, and then find the number of rows to loop over. We will be using the RowCount function. We want the loop through the tickers from i=0 to 11 and we want the ticker= tickers(i). Remember, we need to find out what is the ‘Total Daily Volume’, so let’s create a ‘totalVolume’ variable to go inside the loop and set totalVolume=0. 

![Creating Loops](https://github.com/cynmmarin/Stock-Analysis/blob/2ddcb1572cfdf43d599ed3cac74ef7d723c3c319/Creating_Loops.png)

### Nested Loops

Now, let’s create some nested loops, this is when we have a loop inside another loop. This is necessary to be able to go through the arrays. First, we’re going to find the total volume for the current ticker, then fine the starting price and then ending price for the current ticker. In order to create this code, we will keep in mind conditional statements. Conditionals are expression we write, to check if a condition is true or false. *If* the condition is true, *Then* a block of code will run until it hit the *End If*. When the condition is false, the code will not run. Let’s go over all three conditions. In the following conditionals we evaluate the statement ‘For j=2 to Rowcount’

### Conditional for Total Volume
In order to find the total volume for the current ticker, the condition state that *If* Cells(j,1) equal the ticker, *Then* the totalVolume = totalVolume (which we had set to zero) plus the Cell(j,8). In other words, if the statement is true, we will be looping through the defined data to find out the total volume of transactions each stock had. By going through each ticker on row ‘A’ and adding all the daily volume of transactions on column ‘H’ (which is why we use 8) for that given year. 

![Conditional for Total Volume](https://github.com/cynmmarin/Stock-Analysis/blob/517c4195c41297fd5ae024936924c98a93e307ca/Conditional_for_Total_Volume.png)

### Conditional for Starting and Ending Price

Keeping the same logic as previously explained. If the earlier conditional is false (<>), then we will go an evaluate the ‘Start Price’ and *End If*. Lastly, if the prior conditional is false, we will go ahead and evaluate the ‘Ending Price’ and *End If*. Once all conditions have been coded, we can go ahead and code for where to output the data of the current ticker should go. 

![Conditional for Starting Price](https://github.com/cynmmarin/Stock-Analysis/blob/517c4195c41297fd5ae024936924c98a93e307ca/Conditional_for_Starting_Price.png)

![Conditional for Ending Price](https://github.com/cynmmarin/Stock-Analysis/blob/517c4195c41297fd5ae024936924c98a93e307ca/Conditional_for_Ending_Price.png)

### Formatting and Conditional Formatting

Formatting will help us with the aesthetics of presenting our findings. This will allow for Steve to easily read out our evaluation and present to his parents. A few things we will be doing is formatting our header text to bold, adding, numeric formats to our ‘Total Daily Volume’ column, and adding percentages to our ‘Return’ column. Conditional formatting will let Steve with a glance of an eye, be able to determine the positive and negative returns on a stock. By formatting the cells to green if it’s a positive return and red if it’s a negative return.

###	Creating Buttons

In order for Steve to be able to navigate our analysis without having to rely on VBA we will create some buttons. Using the ‘Button’ function on the ‘Developer’ tab we will go ahead and create a message. This message will read ‘Run Analysis for All Stocks’, it will prompt the question of which stock he wants to see the analysis for, 2017 or 2018 and then will give him our results. The button will make it easier for Steve to understand our findings.

### Measuring Code Performance

Measuring our code performance will allow Steve to understand how long it will take for VBA to perform the code and give him an output of our results. For this we will use the ‘Timer’ function and analyze the ‘starTime’ and the ‘endTime’. Once we run this code, we will have the results be displayed in a message box and thus allow Steve to understand the speed at which VBA process results. We find that for 2017 the code ran in 0.140625 seconds. Meanwhile, for 2018 the code ran 0.691406.   This feature will help give him an understanding of how much longer it will take him to run larger datasets. 

![Measuring Code Performance for 2017](https://github.com/cynmmarin/Stock-Analysis/blob/92dc95764b4ca3c9b1b9c7ce1db6c1f69f6a3fd8/Measuring%20Code%20Performance%202017.png)

![Measuring Code Performance for 2018](https://github.com/cynmmarin/Stock-Analysis/blob/92dc95764b4ca3c9b1b9c7ce1db6c1f69f6a3fd8/Measuring%20Code%20Performance%202018.png)

## Results
### Findings for 2017

Now that we have ran our code and see that it works, lets evaluate our findings. In 2017 we see that the rate of ‘Return’ for the DQ stock is at 199.4% with a ‘Total Daily Volume’ of 35,796,200. Making it the stock with the highest return of all 12 stocks. In this case the overall activity in its trading volume is lowest of all stocks. Meaning is does not get traded as often as a stock such as SPWR with the highest return of all green energy stocks, with a 782,187,000 ‘Total Daily Volume’. What does this say about the stock? There are a few things we can conclude, DQ may not be a well-known stock in the green energy market and thus does not have a high daily volume trading. This makes it a bit risky to invest in a stock given that it is not appealing for investors. It can potentially lack liquidity, making it a great stock to invest short term, but not in the long term.

![VBA Challenge 2017](https://github.com/cynmmarin/Stock-Analysis/blob/832b0ebabbf1add966ecd45956e9b21e50b12944/VBA_Challenge_2017.png)

### Findings for 2018

In 2018 we see that as predicted DQ’s returns go down, reaching a low of -62.6%. This is concerning and reflects that the DQ stock is volatile. Long term investment is highly risky and will result in unprofitable investment. The ‘Total Daily Volume’ decreases to 107,873,900 from the 782,187,00 from 2017. Therefore, we can confidently explain to Steve that DQ is not a stock he should advice his parents to invest all their funds, as they will experience a lost. Alternatively, we can recommend he advices his parent to look into ENPH stock if they wish to continue to invest in green energy. 

![VBA Challenge 2018](https://github.com/cynmmarin/Stock-Analysis/blob/832b0ebabbf1add966ecd45956e9b21e50b12944/VBA_Challenge_2018.png)

We can recommend Steve to look into investing in ENPH, why? In 2017 this stock saw a positive return of 129.5% with a ‘Total Daily Volume’ of 221,772,100 trades. Although is not the highest in 2017, it continues to be profitable in 2018, with a positive ‘Return’ of 81.9% and a ‘Total Daily Volume’ of trades of 607,473,500. We observe that of all 12 stocks, ENPH is one of two stocks that continue having a positive return in in 2018. Therefore, we can encourage Steve to inform his parents of our findings and suggest they consider investing in ENPH and not DQ.

### All Stocks Analysis Refactored

By refactoring our code and altering it to make the VBA script to run faster we find that in 2017 the code ran in 0.1367188 secods, in comparison to the initial 0.140625 seconds. Meanwhile, in 2018 the refactored code ran in 0.1367188, in contrast to the early findings of 0.691406. We observed that by using adding the ‘tickerIndex’ as a variable we make little impact on the 2017 code performance. Although for 2018 there’s a significant increase in performance. Overall refactoring the data seems to be cleaned up our code and make it more efficient.

## Summary
### Advantages of Refactored VBA Script

In general refactoring the code increases the effectiveness of the performance of the code. The advantage of applying this method is that it polishes the internal behavior, without obstructing with the results. Also, it helps us avoid debugging the code. During building the initial code, we ran into stack overflow several times. This was not the issue once the refactoring occurred, the code ran smoothly and took less time to make it run.

### Disadvantages of Refactored VBA Script

Given that the initial refactored code had been built, refactoring the existing code was very time consuming. We had a code that gave us the output desired and although did not ran as fast as the refactored code, it worked. When working on figuring out the procedure that needed to be followed to refactor the code and the alterations that needed to occur, we spent a lot of time. In our case, Steve had not given us a deadline and therefore, having this be a time-consuming task was not a significant disadvantage. However, if we had been given a short deadline, the efficient approach would have been to not construct the initial Stocks Analysis, and instead created the refactored analysis to begin with. 

### Advantages and Disadvantages of the Original and Refactored VBA Script

The original script was easier to build, it applied more simple methods to analyze the performance of the green energy stocks for 2017 and 2018. It also gave us a clear understanding of where the DQ stocks stands in the market. Not to say that the refactored VBA script was not effective. The refactored script rans faster and it’s more efficient, but it was time consuming. 

The disadvantage of the original code, was that there were numerous times that we had to debug our code, normally given to issues with stack overflow. This resulted is having to do tedious work to figure out what was wrong with our code. Alternatively, this was not an issue with the refactored script. Once our variables were properly defined the loops were stablished the code ran smoothy. It became more of an issue of spending time re-working the existing code to make it run fast, and thus be more efficient. In the future, we will attempt to create a refactored code from the start rather that do double work and waste time duplicating our code.

Overall, with substantial confidence given our finding we can advise Steve, to tell his parents to abstain from investing in DQ. DQ is a volatile stock, with short-term profits but long-term loss. Investing all their funds on such volatile stock will result in failing to retain their investment and absence of profits. Alternatively, if his parents wish to abstain from diversifying their portfolio and continue with the desire to invest in green energy, we can recommend he convinces them to invest in ENPH. This stock is steady, with a consistent ‘Total Daily Volume’ of transactions and a positive return in 2017 and 2018.     


