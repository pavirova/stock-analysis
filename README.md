# Stock Analysis Code Refactored

## Overview of Project
I have built a program that helps Steve compare different stocks' performace for 2017 and 2018 in order to help his parents. Once the original code was built and ran successfully, I worked on refacturing the code so that Steve can analyze a larger dataset if he wants, while making sure the code does not take longer to run than the original *AllStocksAnalysis* code. For this project, the goal is to see if refactoring the code improves its performance. The performance of the code will be measured by how long the program takes to run from start to finish. Then, the performance will be compared with that of the previous code, *AllStocksAnalysis.*

## Results

### Difference in Stock Performance for 2017 and 2018
When running the code for years 2017 and 2018, we can see how the different stocks performed by looking at the return column on the "All Stocks Analysis" worksheet. For the year 2017, all of the stocks' performance improved except for the "TERP" stock. In 2017, the "TERP" stock's returns decreased by 7.2%, while in 2018, the "TERP" stock performance decreased by 5%. In 2018, we can see that the performance for 9 more stocks decreased from the beginning of 2018 to end of 2018, putting these stocks' returns as negative. Creating the **For** loops and introducing the array size into each of the arrays we created helped make sure the code ran smoothly through each stock ticker and populated each column successfully.

### Original Script and Refactored Script Execution Times
After running the code for both the original *AllStocksAnalysis* and the refactored *VBA_Challenge* analysis, we can see the difference in running times below:

#### Original *AllStocksAnalysis* Script Execution Times
2017:
![AllStocksAnalysis_Original_2017](/AllStocksAnalysis_Original_2017.png)

2018:
![AllStocksAnalysis_Original_2018](/AllStocksAnalysis_Original_2018.png)

#### Refactored *VBA_Challenge* Script Execution Times
2017:
![VBA_Challenge_2017](/VBA_Challenge_2017.png)

2018:
![VBA_Challenge_2018](/VBA_Challenge_2018.png)

As we can see from the different execution times, it is clear that the refactored code ran faster than the original code. In 2017, the refactored code ran approximately 0.48 seconds faster than the original code. In 2018, the refactored code ran approximately 0.45 seconds faster than the original code. This proves that refacotring our code did help it run faster and more efficiently.

## Summary
Refactoring code has its advantages and disadvantages. Some of the advantages of refactoring code are that it forces you to look at different ways to improve how a problem is solved. Refactoring can help make a code more efficient, improving the code's overall performance by, for example, running faster or using less memory. Additionally, refactoring code can help other users are well as yourself understand the code more clearly. Some disadvantages of refactoring code are that one might be faced with additional errors that they were not facing in the original code. In addition to the errors one could come across, refactoring code can also be time consuming depending on how complex the code is.

For this VBA code in particular, refactoring it took a long time. When attempting the refactor, I came across errors that I did not experience the first time around with the original code. I struggled with understanding why the refactored code was not working, and after different attempts at fixing the errors, I realized that I had my arrays as singular, which is why the code wasn't running through each ticker and outputting the results on the worksheet. A pro of having refactured the code however, is that I learned to fix issues that I did not know how to fix prior to the project, and was able to improve the code's overall performance, as can be seen from the difference in execution times for both years.
