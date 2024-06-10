README
Natalie Raver-Goldsby

Module 2 
I first establish a worksheet For Each statement to ensure that VBA cycles through each worksheet in a workbook. 

Then I define variables that I use later in loops and calculations. 

I place headers. 

I create a loop that moves through each ticker and stock. 

Then I create an If Then Else statement. The if statement condition tests if the cells below row is the same as the one currently in use. As long as that condition is true, VBA sums the opening price, closing price, and total stock. If that condition is false, it calculates the quarterly change, percent change, and total stock and assigns those values as well as the associated ticker to the output table in ssrow. Then the If statement ends. 

The next If Then Else statement calculates the greatest price increase and the greatest price decrease and assigns it to the correct output cells. The greatest price increase assumes that it is the largest. If it is smaller than the next quarterly change, it then assumes that value and tests this against all other quarterly change values. Once the greatest price increase is found to be the largest than any other quarterly change value, its assigned to the correct cell and formatted under Else. The same is performed for greatest price decrease and total stock. 

I apply conditional formatting at the end using excel formatting. I defined the range as quarterly change in column J at the beginning of the script. This code tests if the values in column J are greater than or less than zero and formats the interior cell color as green and red respectively. 

