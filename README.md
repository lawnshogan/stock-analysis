# stock-analysis
Module 2 for Data Science Bootcamp - VBA Scripting



### Stock Analysis - Deliverable 2
###  **Project Overview**
- Steve needs help analyzing stock data from 2017 and 2018. His parents want to invest in one, Steve wants to make sure they are investing their money based on facts and supporting evidence. He's sent you the data to begin the analysis, which will require the use of VBA scripting. The spreadsheet contains data regarding the ticker, Date, Open, Close, High, Low, Adj Close and Volume.

- ### Stock data and its purpose:
    1. Why are we analyzing this data?
    2. What is the goal and possible outcomes?
    3. What pieces of data can help build toward and obtain our goal(s)?

- Steve wants to find the total daily volume and yearly return for each stock. It's important that Steve can interact with the spreadsheet in order to obtain the data from 2017 or 2018.

### **Analysis**
This was personally a learning curve assignment because I had never used VBA in this way before, so I carefully worked through every step. 

First step was creating a module and starting a subroutine.

Within this subroutine, it's important to activate our spreadsheet for the analysis and format your analysis spreadsheet (titles). The Cell and Range were used to assign values to specific cells.
- I wanted the title in 'A1' to reflect what year was being shown based on what data was active.
- Range("A1").Value = "All Stocks (" + yearValue + ")"

Once this formatted, I moved onto creating an array of all the tickers in column A from each sheet.
- I then used a ROWcount (found online) to loop over this column.

####                                              **The fun begins...**

**1a)** This portion of the code uses Dim to assign variable 0 to tickerIndex as an Integer since it's a whole number.

**1b)** Set three output arrays for volume, starting and ending prices.

**2a)** Create a loop to inilialize the tickerVolumes to zero.

**2b)** Loop over all the rows in the spreadsheet using RowCount

We now want to create a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

**3a)** Increase volume for current ticker

**3b)** Check if the current row is the **first** with tickerIndex. If it is, then assign **Start Price**

**3c)** Do the same as 3b, however for the **last row**, then assign a **Ending Price**

**3d)** Used to increase the tickerIndex by 1 when the next cell ticker doesn't match current cell ticker 

**4)**Fianlly, it's time to Loop through the arrays to output the Ticker, Total Daily Volume, and Return.



With these questions in mind, we can begin filtering our data and set parameters that allow us to draw conclusions. **[Pivot Tables](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)** are a great way of setting parameters and summarizing data in an efficient manner. In addition to this, **[Pivot Charts](https://support.microsoft.com/en-us/office/create-a-pivotchart-c1b1e057-6990-4c38-b52b-8255538e7b1c)** are a great way of visualizing pivot tables, which is important in finding patterns.

I was able to use these tools in my analysis to filter by country & category/subcategory to summarize and visualize the number of failed, canceled and successful. Now that I practiced comparing outcomes to categories/subcategories, it was time to compare outcomes based on the launch date using the same methods, using months on the x-axis.

<p align="center">
  <img src="https://github.com/lawnshogan/kickstarter-analysis/blob/main/Theater_Outcomes_vs_Launch.png" width="700"/>
</p>

From here, Louise needed to know about the outcomes based on goals. Creating a new table, I was able to create ranges and use the =Countifs formula to create a summary of outcomes based on pledged dollars for 'play' related Kickstarter Campaigns.

<p align="center">
  <img src="https://github.com/lawnshogan/kickstarter-analysis/blob/main/Outcomes_vs_Goals.png" width="700"/>
</p>

### **Challenges**
The first thing that I will say is every challenge I encountered was fixed by a Google Search. I'm realizing how important the context is when you are searching for answers. This was very helpful and helped me learn. 

I struggled at first with Pivot tables, however they quickly became easier for me after some practice and knowing how to look at the data to make sure you are answering the correct questions.

I especially enjoyed learning about the different Excel formulas and applying them to the analysis. Excel makes math very easy, as long as you are entering in your code correctly!

### **Results**
Theater Outcomes by Launch Date
- The highest amount of successful campaigns in the Theater category were launched in May.
- The fewest successful campaigns in the Theater category launched in Decemeber.

Outcomes Based on Goals

- The campaigns that failed had had higher goals that could not be met.


It would be interesting to see which of these are still active and profiting, which could be used in another analysis in itself. 

I think it's important to look at the Percent funded and Average Donation size as well. A pivot table could be created to show average donation size for successful vs failed campaigns. I believe it would also be important to show the average percent of funding for successful campaigns.

I noticed there is a 'Staff Pick' column in the spreadsheet as well. You could use this to not only filter out successful campaigns, but to go even further and only include those that were picked by staff.
