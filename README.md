# stock-analysis
## Project Purpose:
For earlier stock analysis, macro was prepared using VBA to analyze volume and return for the dataset of 12 stocks in the year 2017 and 2018. But it was apparent that with the existing code, there were very high chances of encountering inefficiencies when the dataset would have been updated to all the stocks being traded in the stock market. 

To overcome the same, the purpose of this challenge was to refactor the stock-analysis VBA code to collect same information on the same dataset as before and make the script run faster to improve performance and scalability of the code enabling the user to quickly analyze more data in terms of years or more stocks.

## Results:
- The 2017 and 18 tabs respectively hold all the data for the stocks.
- DQ Analysis tab has initial analysis only on “DQ” stock that had gathered initial interest of Steve’s parents.
- AllStocksAnalysis and yearValue Analysis has outputs as below:
o	yearValueAnalysis tab holds the initial analysis of Stocks based on the year selection but the code is NOT refactored.
o	AllStocksAnalysis tab holds the same analysis on the same data set of stocks as in yearValueAnalysis tab; and includes the enhancement of refactored code that has boosted both the efficiency and the performance. 
It is to be noted that the results generated on both the tabs are exactly the same but the run time of the code has significantly reduced. 

**Return and Volume analysis on the 12 stocks:**
- In 2017, 11/12 stocks generated positive return from an approximate range of over 5% to about 200%.
- In the year 2018, only two stocks generated positive returns between 81%-84%. Two stocks, positive for two consecutive years include RUN and ENPH. Negative returns on the other 10 stocks was observed to be between the range of – (3% - 63%).

**Elapsed Time for Running Code**:
While refactoring, loop was done to loop through all the data one time.
**Refactoring Status (Yes/No)	Year	Time Elapsed**
- No	2017	7.953125
- Yes	2017	0.109375
- No	2018	3.523438
- Yes	2018	0.109375

The elapsed time for both the years optimized to about 0.109375 which is a great boost from the non-refactored code, thus proving that code quality has been refactored and optimized, significantly improving the overall code performance

## Summary:
## Refactoring Code Advantages 
- Maintainability: To make it easy to enhance and maintain in the future.
- Removing Bad Smell: It prevents many future defects. Code Size is reduced. Confused coding is properly restructured.
- Easier to Understand: Understandable code when you come back to look or refactor further in a few weeks/months, plus extensible code.
- Improved Performance: Makes the code more efficient.
- Testing Support: lead to more unit testable code.
- Debugging Optimization: Clean code makes debugging faster.

## Refactoring CodeDisadvantages:
- Easy Duplication: The disadvantages of refactoring our code are that codes can be easily duplicated.
- Increase in development time: it is possible to underestimate the amount of time for refactoring and end up working on it longer than you planned. And may not be safe. 
- Increase Cost: More time means more cost being spent on a deliverable.
- Delay in Overall Project: If re-factoring is taking more time, it may impact other deliverables for the project.

## Refactoring the original VBA script
### Pros:
One of the major advantages of refactoring code in VBA script is that one can use as much as of the original code as they prefer and can put your new code side by side with old code using different modules and run them in parallel to see results and debug. It is especially helpful in cases of dynamic datasets and analyzing the same data at later point of time.
Another huge benefit is a a decrease in macro run time.
### Cons:
Upon lack of a strong understanding of the syntax and the logic, one will struggle to refactor code and may waste lot of time in delivering measurable results.
Upon doing hit and trials, one may not get the desired results and end up messing the non-refactored code as well.

