<p align="center">
    Stock-Analysis (Delivarable 2)
</p>

<p align="center">
    Module 2 for Data Science Bootcamp - VBA Scripting
</p>



###  **Project Overview**
- Steve needs help analyzing stock data from 2017 and 2018. His parents want to invest in one, Steve wants to make sure they are investing their money based on facts and supporting evidence. He's sent you the data to begin the analysis, which will require the use of VBA scripting. The spreadsheet contains data regarding the ticker, Date, Open, Close, High, Low, Adj Close and Volume.

- ### Stock data and its purpose:
    1. Why are we analyzing this data?
    2. What is the goal and possible outcomes?
    3. What pieces of data can help build toward and obtain our goal(s)?

- Steve wants to find the total daily volume and yearly return for each stock. It's important that Steve can interact with the spreadsheet in order to obtain the data from 2017 or 2018. His families financial future is in my hands, so I better create an accurate analysis.

## **Analysis**
This was personally a learning curve assignment because I had never used VBA in this way before, so I carefully worked through every step. 

First step was creating a module and starting a subroutine.

Within this subroutine, it's important to activate our spreadsheet for the analysis and format your analysis spreadsheet (titles). The Cell and Range were used to assign values to specific cells.
- I wanted the title in 'A1' to reflect what year was being shown based on what data was active.
- Range("A1").Value = "All Stocks (" + yearValue + ")"

Once this formatted, I moved onto creating an array of all the tickers in column A from each sheet.
- I then used a ROWcount (found online) to loop over this column.

## **The fun begins...**

**1a)** This portion of the code uses Dim to assign variable 0 to tickerIndex as an Integer since it's a whole number.

**1b)** Set three output arrays for volume, starting and ending prices.

**2a)** Create a loop to inilialize the tickerVolumes to zero.

**2b)** Loop over all the rows in the spreadsheet using RowCount

We now want to create a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

**3a)** Increase volume for current ticker

**3b)** Check if the current row is the **first** with tickerIndex. If it is, then assign **Start Price**

**3c)** Do the same as 3b, however for the **last row**, then assign a **Ending Price**

**3d)** Used to increase the tickerIndex by 1 when the next cell ticker doesn't match current cell ticker 

**4)** Finally, it's time to Loop through the arrays to output the Ticker, Total Daily Volume, and Return.

   <p align="center">
   Final Code
</p>

<p align="center">
  <img src="https://github.com/lawnshogan/stock-analysis/blob/main/Code%20Block%20VBA.png" width="700"/>
</p>


With that being said, Steve is not going to care about our code. Steve is looking for cold hard results that are easy to read and interpret. As I said above, his family's financial future is in my hands.
- When the code is entered, a pop-up instructs the user to choose what year they wish to get results for.
- Enter 2017 and the following results appear:

<p align="center">
  <img src="https://github.com/lawnshogan/stock-analysis/blob/main/VBA_Challenge_2017.png" width="700"/>
</p>

- Run the script again and enter 2018 instead:

<p align="center">
  <img src="https://github.com/lawnshogan/stock-analysis/blob/main/VBA_Challenge_2018.png" width="700"/>
</p>

### **Results**

### 2017:
After the code ran for 0.07 seconds, the results for 2017 showed the following:

1. TERP was the only stock which would have produced a negative return (-7.2%).
2. FSLR had the largest daily volume and was 4th in returns. 
3. DQ, ENPH, FSLR & SEDG all had return gains over 100%.

### 2018:
When running the code for 2018 (0.08 seconds runtime), the results told a different story:

1. All stocks produced a negative return, with the exception of ENPH (81.9%) and RUN (84%).
2. ENPH amd RUN also have the top 2 in total daily volume.
3. DQ, FSLR, JKS, SPWR all had losses of over -35%.

Based on these results, and as an extremely unqualified tax specialist, I might personally want to invest in DQ and SEDG based on the difference in their gains/losses. 
- DQ had almost a 200% return in 2017 compared to a -62.6% in 2018. Lets hope for a better 2019. 
- SEDG exceeded expectations in 2017 and gained 184.5% in returns, compared to a -7.8% loss in 2018.

Hopefully his family does not lose all their money.

### **Summary**

After refactoring the code for this assignment, it made me realize that although it can be tedious, frustrating and time consuming, it is more efficent to refactor code that's already been created than to start from scratch. Someone has already created the code you need (most likely). Your job is to find it and make it work with your project. 

I found myself taking code I found on Google and applying it to my module, however I found it very easy to get confused unless you are only working with one part of the code at a time. I found it was easier taking a larger, complete block of code and edit off that, rather than try and piece together multiple sources of code. 

Another disadvantage I found was how many different ways there are to do one task. Sometimes piecing together code can cause conflict and this defenetly set me back at times. 

Regarding the refactored VBA script, I think it was helpful for the pieces before 1a and after 4. Your sheet is already formatted and ready to go. Copying it over is easy and allows you to direct your attention to the other portions of code we haven't tested, or even created yet.
