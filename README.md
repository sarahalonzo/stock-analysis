# **VBA Challenge for Stock Analysis** 

## **Overview of the Project and Purpose of the Analysis**
  &ensp; The VBA Challenge began with an Excel workbookbook with data regarding 12 stock tickers. Client Steve and his parents need to evaluate the performance of their stock and other portfolio stock total
  volumes and returns in order to make investment decisions.

  &ensp; Through use of VBA coding in Excel the analyst can establish subconditionals that contain loops with conditionals and other arguments that produce efficiently accessed 
  and formatted data showing the performance of all 12 stock tickers in 2017 v. 2018. The analyst has to use VBA coding in order to loop through all tickers for their 
  total volume and returns.  Without VBA the task of computing and formatting the large data sets would be long and inefficient.

  &ensp; Code is then refactored to allow for faster results, debugging, and future building and growth of data and code. Client Steve and his parents are able to quickly 
  retrieve data needed on the total volumes and returns for all tickers in 2017 v. 2018 with the use of a macro linked button in the spreadsheet. A clear data button 
  is also programmed in order to clear out the screen should the user have the need.

## **Results of Analysis**

### **Comparison of Stock Performances in 2017 and 2018**
       
  &ensp; Through use of VBA code it is clearly visible that client Steve and his parent's stock DQ has fallen from 199.4% green or positive returns in 2017, into the red
  at -62.6% returns in 2018. Overall, most tickers displayed have fallen into negative returns from 2017 to 2018. 

  &ensp; The two most consistent stocks that have maintained a green or positive status are the second highest average return stock ENPH, and secondly the stock RUN. These are two stocks that Steve and his parents may want to monitor for current or future investment. The data range for stock investment advice is short in order to advise for future investments based on a pattern of performance given a span of two years in the data. Note that the length of the data set is a limitation, assuming that there are more years of data and trends available.

***Please see stock data results in VBA Challenge Workbook:***

 https://github.com/sarahalonzo/stock-analysis/blob/44509f2697f93e40289715a25144bea5aed22e01/VBA_Challenge.zip


### **Execution Times Before and After Refactoring**

  &ensp; Prior to refactoring there were 5 subfactors established within the code. Following refactoring there are 2 subfactors in the code. There is only a second subfactor 
  because coding was reestablished at the end to reprogram the button macros where a second sub was added for the clear data button. 

  &ensp; New code with a loop that contained all conditionals and arguments across tickers was established through a ticker index containing arrays that reduced the time it 
  took to run the code in the "green_stocks" original workbook by almost half of a second.

  &ensp; For 2017 run time was reduced by .3437475 seconds. 2018 refactored run time was reduced by .359375 seconds. Please see a comparison of the run time images from the
  macro below starting with refactored images as compared to the last two original run time images.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/95314378/147019208-88bf82ce-d4b6-4bf3-8e3f-74562283cce5.png)

Resources/VBA_Challenge_2017.png

https://github.com/sarahalonzo/stock-analysis/blob/f6af6a6814765800f7aad337342cc9dcdecb43db/Resources/VBA_Challenge_2017.png


***VBA Challenge 2017***

![image](https://user-images.githubusercontent.com/95314378/147019323-3161c553-b92a-45ea-a4fe-6cb44d46e505.png)

Resources/VBA_Challenge_2018.png

https://github.com/sarahalonzo/stock-analysis/blob/f6af6a6814765800f7aad337342cc9dcdecb43db/Resources/VBA_Challenge_2018.png

***VBA Challenge 2018***

![greenstocks_challenge_ 2017](https://user-images.githubusercontent.com/95314378/147019375-5dae86b6-39cd-474d-901a-0886fb63ac72.png)


Resources/greenstocks_challenge_ 2017.png

https://github.com/sarahalonzo/stock-analysis/blob/f6af6a6814765800f7aad337342cc9dcdecb43db/Resources/greenstocks_challenge_%202017.png

***green_stocks 2017 Original Code Run Time***

![greenstocks_challenge_ 2018](https://user-images.githubusercontent.com/95314378/147019415-8bbdc7ba-e807-4fa9-bd0f-c6d0f553b3b5.png)

Resources/greenstocks_challenge_ 2018.png

https://github.com/sarahalonzo/stock-analysis/blob/f6af6a6814765800f7aad337342cc9dcdecb43db/Resources/greenstocks_challenge_%202018.png

***green_stocks 2018 Original Code Run Time***

## **Summary: Advantages and Disadvantages of Refactoring**

  &ensp; The advantages of refactoring include debugging, faster run times in this particular analysis, cleaner code to read and interpret, simplified allowance of additional data and/or 
  code analysis to be added, edited, or altered in the future through additional nesting. 

  &ensp; The disadvantages of refactoring include adding bugs to your code, disabling previous features or arguments that may have been wanted or needed retention, 
  forgetting to reinstitute items such as buttons that following refactoring may not be valid or perform as prior to refactoring. Refactoring may also confuse
  basic users who are not as familiar with nesting funcitons and arguments within code  and prevent some viewers from following the entire code story. The code story or details may
  be important in seeing how the analyst came to their final conclusion/s and in having the entire picture of how the different analysis was built in steps to discover the
  level of relational data.

    





