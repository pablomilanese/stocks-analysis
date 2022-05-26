# Analysis of the Stock Market
## Background
Steve is a recent graduate in finance. His parents, who are passionate about green energy, have decided to invest in the stock market. They took a keen interest in DAQO New Energy Corp, a company that makes silicon wafers for solar panels. In order to help his parents diversify their portfolio beyond DAQO, he gathered data for various green energy stocks.

To help Steve with his analysis I created macros in VBA. These macros will help Steve analyze the stock market with the click of a button. Moreover I refactored the macros to diminish the run time for each macro.

## Purpose
The purpose of this project was to refactor the macros I created for Steve. By refactoring the macros we guarantee that they will work even when Steve runs thousands of rows of information. Although the original macro does not take long to run it is probable that it will slow the more data Steve feeds it.

## Results

##### FIRST STEP · PREPARATION

I wrote out the different steps that were needed in order to keep my coding clear and organized. By doing so I am helping the next person that looks at my coding understand what were the steps I took and how each section works.

The outline looks as follows:

```
'1a) Create a ticker Index
'1b) Create three output arrays
'2a) Create a for loop to initialize the tickerVolumes to zero.
'2b) Loop over all the rows in the spreadsheet.
'3a) Increase volume for current ticker
'3b) Check if the current row is the first row with the selected tickerIndex.
'3c) check if the current row is the last row with the selected ticker
'3d) Increase the tickerIndex.
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
```
> It is important to add a single quotation mark in VBA in order to keep this text as a comment and not part of the code.

##### SECOND STEP · CODING

Once I had my outline written out I began coding each step:

<img width="703" alt="codingprocess" src="https://user-images.githubusercontent.com/105120795/170198451-885358dd-239a-4a14-bb07-6eb86f48ab6f.png">

###### SIDENOTE
Although writting readable code is important for the sake of future analysis, it is also important in some instances to make sure our macros work.
I had originally ran into some issues when I had organized all my code with no tabbing. The only formatting I used was line breaks between each task.
After analyzing the code in Visual Studio I realized that some statements ('For Next' for example) were getting nested under the wrong counters.
I have not had much time to further look into this but the reality is that tabbing my code properly helped debug my code. This strengthens the importance of writting readable code and formatting it properly.

##### THIRD STEP · TESTING

Once I had my refactored code working I tested both subroutines: AllStockAnalysis (original) and AllStockAnalysisRefactored (refactored code).
The purpose of these tests was to measure code performance. If done correctly, the refactored code should run faster. Although the data that Steve has so far is not that large, future data documents could slow down our macros, which would be counterintuitive. As stated on the Purpose section of this analysis, the goal is to help Steve analyze large amounts of data in the shortest time possible.

These were my results:

Original subroutine times

<img width="422" alt="VBA_Challenge_Original_2018" src="https://user-images.githubusercontent.com/105120795/170205656-353161b8-5467-41d5-bb60-03bd90083818.png">

>0.60 seconds aproximately.

<img width="420" alt="VBA_Challenge_Original_2017" src="https://user-images.githubusercontent.com/105120795/170205616-5fb69cfb-9a22-4358-a6a8-54004e0b8ac5.png">

>0.59 seconds aproximately.

Refactored subroutine times

<img width="421" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105120795/170205713-3092402d-69a0-43a8-995c-56c9a8f111ed.png">

>0.11 seconds aproximately.

<img width="421" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105120795/170205779-1dba0b11-9472-4263-a2d8-1d6c7d9c4bc8.png">

>0.11 seconds aproximately.

How the different time measurements compare side-by-side:

<img width="474" alt="Screen Shot 2022-05-25 at 2 34 33 AM" src="https://user-images.githubusercontent.com/105120795/170206983-6624f836-e318-4321-8966-ef941080efb4.png">

Even though in both macros the results are returned in less than a second, it is clear that the refactored code shorted the run time.

## Summary

### Advantages of Refactoring Code

### Disadvantages of Refactoring Code

### Advantages of the Original Code

### Disadvantages of the Refactored VBA Script











