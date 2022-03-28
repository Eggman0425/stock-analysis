# VBA Challenge
Refactoring the solution from Module 2
# Overview
The purpose of this analysis is to take existing information of the stock market and analyze the past few years of certain stocks with the click of a button.  The button will display each stock ticker, the stock’s total daily volume, as well as color coding the return percentage. By creating this button, Steve’s parents will be able to see this information displayed cleanly in under a second.
# Results
The stock performance in 2017 was significantly more successful in returns than in 2018, with only the return for TERP in the negative. While in 2018 most returns were in the negative, only ENPH and RUN returns were successful. 
In the original code the execution times were approximately one second, but with the refactored code the run times decreased to less than 0.2 seconds. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100382595/160311132-87f94c4b-318e-4779-b0a5-fa078b1d09af.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/100382595/160311136-0c4b66bc-9eec-426a-ae2c-67d205d380c4.png)

This is large part due to the refactored code increasing the volume for the current ticker and increasing the ticker index in the second for loop. This helped add the volumes for each ticker more quickly, as well as increase the return for each ticker.

![Stock_Analysis_Code](https://user-images.githubusercontent.com/100382595/160311200-da9c48c2-cddb-4709-b377-f90484f2927f.png)

# Summary
The advantages of refactoring code are improving existing code into more easily understandable code, as well as developing a more efficient route to the same answer. The disadvantage of refactoring code is trying to achieve the same result with a slightly different equation, as I ran into some issues tracing back my steps in the original code. With the original VBA script, the code was tedious and long, but the refactored code made it more concise along with increasing the speed of the macro. 
