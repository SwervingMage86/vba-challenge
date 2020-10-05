# vba-challenge

# Building Stock Analysis Code

# Step 1:

I knew the first thing I needed to do to get this code off the ground was find the last row of the entire data set.
I did this by using the rows.count function to go to the end of the table and count the rows up to get the entire array of the data set.
I then started an For Statement to loop through the entire array

# Step 2:

Before I could do anything else, I needed to find where the last row of each ticker so that I can pull the statistical data I needed for each ticker.
I did this by building and If statement to search for the row where the ticker values no longer match and pulled my end values (end price and end volume)
As it turns out I didn't end up needing the end volume of the ticker but left it in to show my thought process.
I then realized I needed to define an open price for my calculations so I added that in before my For Statement as it was just an intial cell volume I didn't need it in my loop.

# Step 3:

I added in my calculations to the code and placed each result in a summary table.
I created the summary table by defining the row number; as ST_Row; I wanted the table to start in and when placing the values and then using Cells().value to define which column I wanted the values to fall in.
To go to next row of summary table I used ST_Row=ST_Row+1 so the code would know to fill into the next row. 
I also added titles to my summary table at this time.

# Step 4:

When running this part of the code, I didn't have a way to change the open price to the next tickers opening price and was returning wrong numbers.
I found this out by setting open price of each ticker to fall into column 10 and end price to fall into column 11
This showed that all my opening prices were the same.  I fixed this by adding Open_Price=Cells(i+1,3).value at the end of my sub.

# Step 5:

I was still unsure at this time how to calculate total volume so I left this and started running my sub on other worksheets to test it.
It worked on all the alphabetical_testing sheets except for worksheet "P"
I knew I was on the right track so I opted to solve this later.

# Step 6:

I added in conditional formatting, starting with changing the percent change column to a percentage format using the .NumberFormat function.
I then created an If/Else Statement to change colors of yearly change column to green if above zero, else red if below zero.

# Step 7:

I created a way to determine total volume of a ticker by adding in and Else Statement for the orginal if statement, basically telling the loop to tally the volume column as long as the ticker symbols matched
I also realized that I needed to define a starting number for the total volume, so before my inital if statement, I defined total volume=0
I sorted my data to one ticker and manually calculated a couple ticker volumes to compare with my sub calculation and they didn't match
I realized that I wasn't adding the final volume value, so had to insert a formula into my if statement to ensure this would happen

# Step 8:

Code was working on everything but "P" worksheet, so I went ahead and moved my sub into the Multi Year workbook to see how it would run
It worked until it hit the P's in the first year I ran it in so I knew I had to find a solution to the problem I was having
I noticed that both the 2014 year data and the "P" worksheet (from 2014) both stopped at ticker "PLD"
I sorted to the next ticker and realized all values for that ticker were 0 and determined that the error was a div/0 error
I solved this by inserting nested If Statement to set the value for open price to equal Null if open price equaled 0
Then added an Else Statement to perform the calculation for the percentage change if open price was not 0 and then place that value in a cell
The code was then running as needed.

# Bonus

# Step 1:

First thing I did for the bonus work was add titles for another table

# Step 2:

I found the greatest increase and decrease by using a worksheetfunction.max and worksheetfunction.min, respectively, in the entire column of percentage change from my summary table.
I did the same for greatest volume by using the max funtion for the volume column
I then used the match function to place the ticker value in the preceding column, however at first it was just placing the number of the row the value it was matching was in.
To solve this I assigned the row number it was returning to another variable, then used that variable in a Cell().value funtion to grab the string value of the ticker in that row and return it into the column I needed it in.

# Step 3:

I formated the percentage change cells using .NumberFormat and printed all my values in the correct cells.

# Step 4:

To loop through each worksheet in the work book, I dimmed ws as Worksheet
Then I used a For Each statement to loop through each worksheet.
I had to make sure to go in and put ws. in front of each cell or range value in the code to make sure that it was pulling from the active worksheet
At the end of the code I added ws.Activate to activate to the next worksheet