# VBA_Challenge
## Overview of Project

### Purpose

This project is about refactoring a visual basic code. Changes to this code were done to loop through all the data at one time. After having run the program, I will present the analysis and the results. In the end, I will present the advantages and disadvantages of refactoring a code. 

## Analysis

In the first part of the code, we can find the startTime and endTime variables. It is important to declare them so we have a place to store the timer. We have also a yearValue variable that will take the input from the user, you can write 2017 or 2018 and it will run the data that matches that year. 

After that, we activate that worksheet from which the next part of the code will take action. Then, we add titles and subtitles for each column. Then have the tickers array, this section is very useful, the code will loop through that and compare data. 

### ![Code1](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Code1.png)

In the next section, we declare the next variables that are going to be used: tickerIndex, the i iterator, tickerVolumes, and Startind and Ending Prices. 

![Code2](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Code2.png)

In this section the code is looping into each ticker and for each ticker we have a second loop. The second loop will be adding the volumes of that ticker into a variable and setting the tickerStarting and tickerEnding Prices. The second loop will be executed for 

![Code3](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Code3.png)

 Section number 4, will add the results to the "All Stocks Analysis" Sheet. Then, we have the formatting section in which we are adding static formatting for titles and subtitles and conditional formatting for column 3. 

![Code4](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Code4.png)

## Results

### All Stocks 2017

We got this chart with just one click. We can see each ticker's total daily volume and the percentage of return for that year. Thanks to the conditional formatting, it is easy to read the results and we can see that the best ticker of 2017 in volume was SPWR and in the percentage of return was DQ. Even though the TERP ticker had a good volume, it got a negative percentage in return. 

![Results2017](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Results2017.png)

### All Stocks 2018

For all sticks 2018, we can see that the numbers are very different from 2017. The highest volume this time was from ENPH and the best percentage of return was from RUN ticker. Even though some volumes improve, the percentage of return went negative almost for all the tickers during this year.  

![Results2018](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/Results2018.png)

We had the data for one year analyzed in less than 0.3 seconds. We got almost the same amount of time for the 2017 analysis and the 2018 analysis. 

![VBA_Challenge_2017](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/VBA_Challenge_2017.png)

![VBA_Challenge2018](/Users/dylanmontemayor/Documents/Professional/Bootcamptec/Excel VBA/Challenge2/Resources/VBA_Challenge2018.png)

## Summary

### Advantages and disadvantages

#### Refactoring code in general

The main advantage of refactoring a code is that you can improve it in many ways: you can save storage, improve the processing time, and reduce the amount of code. One of the main disadvantages of refactoring could be changing a lot of the code and finding out that it doesn't run and spending too much time debugging. 

#### The original and refactored VBA script

The main advantage is that after having the code refactored, you can save a lot of time in running it, especially if you add a lot of data. If you don't have much data, you probably won't notice a lot of the difference. 
