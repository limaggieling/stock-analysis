# stock-analysis
 using vba, create macros to automate stock analysis data

## Overview of Project

**Purpose**
The purpose of this analysis was to utilize VBA to code a macro that would automate the process of analyzing different stocks chosen by Steve's parents. Initially, we only analyzed the DQ stock by reading the 2017 and 2018 sheets to extract the closing prices at the start of the year compared to the end of the year. Also, we wanted to find the total volume traded which we assume to be an indicator of the success of the stock. Then, we wanted to apply the same analysis on all of the stocks. Lastly, we wanted to refine our code to run efficiently.   

**Background**
Coding through VBA allows us to create macros that can automate trivial, repetitive tasks in Microsoft Office Excel. By creating macros, we can efficiently screen through data and output information that we can use to make inferences. For example, we can run a quick analysis that will spit out the percent return and total volume in less than one fifth of a second. We used variables with different data types to fit the values we needed to store, if-then statements, For Loops, comparison operators, indexes to access data within arrays, nested loops, conditional formatting, and numeric formatting to build our macro for Steve. We practiced refactoring our code to improve its runtime. Learning to script streamlined code is similar to revising a written report, it should be concise, simple, and effective in conveying its purpose.

## Analysis

**Results** 
We wanted our code to be capable of reading through the provided data from the years 2017 and 2018. We implemented the command InputBox to allow the user of our macro to type in the year they wanted to analyze. The command allows for a prompt to be given and for the user input to be saved as a variable in the form of a string. We saved the users' input under "yearValue".

![inputbox_code](Resources/inputbox_code.png)

We set up our output sheet to have 3 columns; ticker, total volume, and return, that would save our calculated values to their respective stock ticker name. We took advantage of using an array to store the ticker data and accessed the specific stock ticker by its index. Then, VBA was able to readby looping through each row and updating the output columns through the use of if-then statements so long as the ticker value matched the stock ticker we wanted to analyze.

![for_loop_if_then_statements_code](Resources/for_loop_if_then_statements_code.PNG)

At the beginning of our code, we also used the timer function in VBA to save the start time after getting the user's input year value and at the end of the code, after the calculations were saved in the output sheet. This gave us the total elapsed time so that we could quantify how efficient our code was. 

our completed code looked like this:
    
    Sub AllStocksAnalysis()
    'initialize startTime and endTime
    Dim startTime As Single
    Dim endTime As Single
    
    'activate new sheet all stocks analysis as output sheet
    Worksheets("All Stocks Analysis").Activate
    
    'have user input the year for stock analysis, give prompt and title
    yearValue = InputBox("What year would you like to run the analysis on?", "Year for Stock Analysis")
    
    'start timer function on excel after user inputs year to analyze
    startTime = Timer
    
    'sets value in cell a1 with title
    Range("a1").Value = "All Stocks (" + yearValue + ")"
    'changed "All Stocks (2018)" to "All Stocks(" + yearValue + ")"
    
    'sets the headers in columns A,B,C on row 3
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "% Return"
    
    'to run analysis on all stocks, create a program flow that loops through all the tickers
    'obviously we can just copy and paste the code we used for DQ over and over for all tickers
    'but to be more concise and precise with coding its better to be DRY: Don't Repeat Yourself
    'Let us use a For loop and an Array
    
    'Arrays: aka lists
    'they hold an arbitraty number of variables of the same data type
    'each variable has an **index** and the array is named
    'Index: is the position of the variable in the array, start at zero
        'ex:index of 1 is the second position in the array
    'arrays are initialized with the dim keyword too but you need to
    '1. insert a number in parentheses to represent the number of elements
    '2. specify the type of variable for each element in the array
        'ex: Dim tickers(11) as string : Dim arrayName(number of elements) as dataType
 
    Dim tickers(12) As String
    
    'assign each element, total of 12 stocks, sorry no shortcuts here
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   
'2.3.3 Reusing Code
' note that you shouldn't initialize variables inside loops as it will confuse VBA
' Make a plan
    'Our new macro should do the following:
'1. Format the output sheet on the "All Stocks Analysis" worksheet. (complete)
'2. Initialize an array of all tickers. (complete)
'3. Prepare for the analysis of tickers.
        'Initialize variables for the starting price and ending price.
        Dim startingPrice As Double
        Dim endingPrice As Double
        
        'Activate the data worksheet. changed 2018 to yearValue variable
        Worksheets(yearValue).Activate
        
        'Find the number of rows to loop over.
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
'4. Loop through the tickers.
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
    
'5. Loop through rows in the data. changed 2018 to yearValue variable
            Worksheets(yearValue).Activate
            For j = 2 To RowCount
        
        'Find the total volume for the current ticker.
                If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                End If
        
        'Find the starting price for the current ticker.
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
                End If
        
        'Find the ending price for the current ticker.
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
                End If
    
            Next j
'6. Output the data for the current ticker.
   
            Worksheets("all stocks analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
   
        Next i
        
'end timer function at the end of the analysis
endTime = Timer

'show elapsed time in messagebox
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

These were the results of from our initial code:

![all_stocks_analysis_elapseruntime_2017](Resources/all_stocks_analysis_elapseruntime_2017.PNG)

![all_stocks_analysis_elapseruntime_2018](Resources/all_stocks_analysis_elapseruntime_2018.PNG)

As you can see, it took almost 1 second to run the analysis.

While refactoring our code, we realized we could condense 


Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

**Summary**
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?



