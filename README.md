# Tickers analysis

## Overview of Project
Steve parents are pationed to alternative enegies. They invest all their money in the DAQO New Energy Corp (DQ). Steve need to look into the DQ stocks but he also wants to diversify the founds of his parents. So he wants to analys some data about green enegy companies.

### Purpose
We want to check the tendency of the stocks of DQ to determine if it's a good option to invest. Also we want to find some other options to diversify the founds of Steve parents.

## Results
![Total Daily Volume All Stocks 2017](https://user-images.githubusercontent.com/88845919/134816251-c756662c-a89b-4eb7-aa82-92ed30d4b5f9.png)
Checking the total daily volume results for 2017, we can see that DQ is the company with less volume at this year. So it's a small company that it's starting porbably. We can spect to increase their volume and to improve the return if it starts growing. 

![Total Daily Volume All Stocks 2018](https://user-images.githubusercontent.com/88845919/134816262-cf6a1415-e08f-4638-a05e-b5dd302252ea.png)
We can see that in 2018 DQ actually increase their volume, but it wasn´t the company with better performance. Actually the company that increases more their volume was ENPH. This make this company a better option, but we need to check the return % to determine the best company to invest.

![Return All Stocks 2017](https://user-images.githubusercontent.com/88845919/134816263-de484a8e-ba08-40e3-b487-3e1dfc8741ca.png)
We can see in the return for 2017 graphic that actually the return of DQ at that year was very good. DQ was the company with the higher return %.

![Return All Stocks 2018](https://user-images.githubusercontent.com/88845919/134816270-eb412144-f649-4854-90b1-88fb44576d30.png)
But what happen at 2018. The return of DQ was negative -62.6%. So even their increase the volumen, the return was very bad. 

![Volume and Return 2017 vs 2018](https://user-images.githubusercontent.com/88845919/134816282-515e86ab-db43-4117-b018-d093758f6163.png)
Now looking at all this information, we saw that ENPH increses a lot their volume, but the return % dreceases at 2018 compare to 2017. We can notice that the RUN company has a posotive increase on the volume and return. So, this should be the better company to invest. If we want to diversify, RUN and ENPH could be good options.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/88845919/134816551-b6839132-676f-4838-a779-7340ff1f7f5c.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/88845919/134816545-877dd865-153c-4e34-9556-6502922a6391.png)
Talking about the refactoring of the code, we notice a improvement on the time that the program runs. From 1.81 to 1.65 sec. on the 2017 information, and 1.74 to 1.54 sec. at 2018 analysis. Mainly because we can make operations and storage reparetly. Also, using variables make much easier to manipulate the program and functions and to update the script if we need to change something.

For this first code, we make the operation of the volume for each of the tickers, and when we finish we print the value on the cells. So we interact with both worksheets at the same time. This make this script slower.1

###Firts code
Sub AllStocksAnalysis()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    Sheets(yearValue).Activate

       startTime = Timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("AllStocksAnalysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("AllStocksAnalysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   Call formatAllStocksAnalysisTable

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

We improve the code adding some variables and arrays to change information faster, only interact with one worksheet at the time. We fix a time to get all the information at once. And then we print the info at the cells. At the end we gave some format. The actions are divided, not all at the same time.

###Refactorung code
Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As LongLong
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For x = 1 To 12

        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
                
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

            End If
                
            'End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            'End If
            End If
            
        Next i
        
     Next x
        
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
            
        Worksheets("AllStocksAnalysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
            
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


## Summary

What are the advantages or disadvantages of refactoring code?
The advantage is that we can make our program faster, it requires more time to do it this way, thinking in different ways to optimize it. But at the end if we have a very large code, we can manage easily and we improve a lot of aspects. One disadvantage is that at some point we are limited on this improvement. Also using arrays makes much easier to manipulate information. We can storage and work with it later.

How do these pros and cons apply to refactoring the original VBA script?
At this scripts, probably we can not notice the diference in time, but it actually makes more easier to make changes in the code. We only need to identify the variables and update the performance if we want to expand the dataset.
