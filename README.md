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

The main advantages of refactoring code are making it more efficient and easier to understand. Although we might have an initial outline for our code we may overlook other ways to get the same results in less steps or making the code run in shorter periods of time. Because our initial purpose with any macro is to make it work we may take a longer route to get there. Refactoring code that already works allows us to focus on what it does and how can we streamline it. It also gives us a second chance to make sure the code will make sense to anybody that has to analyze it. Just like an essay, our work may diverge from the original outline and become harder to understand. 

### Disadvantages of Refactoring Code

There are some disadvantages of refactoring code:

It can be time consuming - refactoring code can be counterproductive if the programer does not have a lot of time to work. The pressure of deadlines can make it harder for them to fully understand how the original code works and find where it can be improved.

It can have diminishing returns - after a programer refactors code they might come to the realization that the changes they have done do not have a huge impact on the performance or readability of the code. When confronting the challenge of refactoring code it is important to keep track of how long the process it taking. If the refactored code does not improve the performance of the original macros, the programmer has to question how helpful his re-work has been so far and if it's worth working on it further.

It can over-complicate code that was already working properly - while trying to refactor code the programer may end up making code harder to understand or may end up developing errors that the original code did not present.


### Advantages of the Original Code

The main advantage of the original code if that it fullfills it's original purpose: make it easier for Steve to analyze the stock market at the click of a button. The code already has conditional formatting to make it easier for the user to navigate results and the macro already has comments explaining how it works. In the worst case scenario Steve has to analyze a large amount of data that may end up taking minutes instead of seconds. Even then, the only thing Steve has to do is press a button and our macro will do the rest.

### Disadvantages of the Refactored VBA Script

One of the disadvantages of the refactored VBA script is the additional time it took to complete the project. Although the refactored code did run in a shorter amount of time, the difference was too small to justify the time spent on it. It would be interesting to run larger amounts of data to see how both macros perform. Although we can mention to Steve that our new work runs faster the results are too close to each other to make an impression on him. Refactoring is a process that will mostly interest other programers. If we have to present a project to an individual that has little understanding of the intricacies of programming they may see our additional work as a negative instead of a positive.











