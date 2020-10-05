# Stocks Analysis Refactoring

## Overview of Project

### Refactor Module 2 solution code to loop through all the data one time in order to collect the same information done in this module.  Then, determine whether refactoring the code succesfully made the VBA script run faster.

## Results

- Through the analysis of the output and performance we can get the following conclusios:

1.	Stock Performance: 
    - Although DQ has had a a stellar year in 2017 (+199.4%), it was the worst performer in 2018 (-62.6%).  So we can conclude this stock has a very high Beta to the stock markets, making it highly riscky.
    
    	- 2017
    	![2017](/VBA_Challenge_2017.png)
    	- 2018
    	![2018](/VBA_Challenge_2018.png)
    
2.	Execution Times:
	- There was a substantial improvement in the refcatured code compared to the original script, the code is runnig more than 4x faster. 
		
		- 2017
			- Before
			![Old2017](/No_refactoring_2017.png)
			- After
			![New2017](/VBA_Challenge_2017.png)
	
		- 2018
			- Before	
			![Old2018](/No_refactoring_2018.png)
			- After
			![New2018](/VBA_Challenge_2018.png)
## Summary

- We can imply some advantages and disavantages of refactoring code:
	- Advantages:
		- Improve readbility;
		- Reduce complexity;
		- Improve source code maintanability;
		- Create a simpler, cleaner or more expressive internal architeture
		- Improve extensibility
		- Improve performance
	- Disavantages:
		- Time consuming
		
- These pros and cons applied to refactoring the original VBA script resumes in, although consuming some time in the process, it brought much more advantages in having a cleaner and more flexible code, which could be used in a much higher number of stocks, and improving the perdormance in more than 4 times.

