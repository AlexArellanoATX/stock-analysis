#   **Module 2 Challenge: Refactoring VBA Code**


##  Overview of Project: 
This project demonstrates how refactoring existing code can improve and optimize its execution and reduce the time it takes to execute by using more efficient code.

### Explain the purpose of this analysis.
The purpose of our project is to make the VBA code we wrote to analyze our stock volume and price data more efficient and run faster so that it can be used to analyze much larger stock data sets, without straining the computing resources or time limitations. After writing initial VBA code where we used two "for" loops and several conditional statements to add up the trading volume for each stock by year, we refactored the VBA script to produce the same output in less time. The refactored code added 3 arrays to the original code and used an index variable to iterate through the stock dataset once instead of multiple times and output the annual trading volume and annual stock performance for each stock. The index variable is used to access the corresponding data for each different stock in the stock ticker index and the loop collects the required output data more efficiently. This refactoring of the code makes it possible to run the analysis faster, even though the output results and functionality of the code do not change.


##  Results: 
### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

While the original VBA code we wrote works to show the performance and trading volume of the stocks in our dataset, using 2 loops requires the dataset to be looped through many more times than necessary. Below is the relevant portion of the original VBA code:

 ![OriginalVBACode_Loops&conditional_Snippet](./Additional_Resources/Original_VBAScript.png)

The refactored VBA code has the exact same funtionality, but uses 3 additional arrays that we created to identify and store the necessary data. We used an index variable named "(tickerIndex)" to access the corresponding volume and closing price data. 

 ![RefactoredVBACode_Loops&conditional_Snippet](./Additional_Resources/Refactored_VBAScript_Loops&Conditionals.png)
 
 The final loop we wrote in the refactored VBA code outputs the volume and closing price data from the new arrays we created, and this process requires less loops than the original code, which makes it faster and more efficient.

 ![RefactoredVBACode_OutputDataFrom3NewArrays_Snippet](./Additional_Resources/Refactored_VBAScript_OutputDataFrom3NewArrays.png)

 This refactored code is longer and more complex to write, but serves the purpose of reducing the amount of time required to run the analysis. 
 
The original code we wrote to analyze the stock dataset executed in between 0.53 and 0.56 seconds, for the 2017 and 2018 stock dataset:

 ![OriginalCode_2017Perf](./Additional_Resources/DQAnalysis_Sub_Performance2017.png)



 ![OriginalCode_2018Performance](./Additional_Resources/DQAnalysis_Sub_Performance2018.png)

 
 


 while the refactored VBA code executed in slightly over 0.10 seconds, for both years.  


 ![RefactoredCodePerformance2017](./Resources/VBA_Challenge_2017.png)
 ![RefactoredCodePerformance2018](./Resources/VBA_Challenge_2018.png)

 
  This is significant performance improvement: a greater than 80% drop in execution time.
Using arrays and an index variable to find and store the daily trading volume and closing stock prices for each stock, improves the efficiency of the operations and this makes it possible to reduce the time to run.



##  Summary: 

#### Advantages of refactoring code include:
* reducing the time and system resources required to run code
* identifying, reducing and avoiding bugs that could be time consuming and expensive to fix later 
* making it easier for multiple people to work with and undesrtand the code, when they are maintaining or adding to the code 
* understanding the data, code and functionality it provides better can make development and analysis teams more efficient


### The disadvantages of refactoring code include:
* the time and effort it takes to identify, write, test and produce the refactored code
* opportunity cost of not spending time working on something that generates revenue or creates new code

##### How do these pros and cons apply to refactoring the original VBA script?
