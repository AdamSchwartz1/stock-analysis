# Stock Analysis with VBA

##Overview of Project

###Purpose

The purpose of this project was to analyze stock data and gather trading volumes and returns over the given year. I created a VBA script to automate this process for Steve. I also added in some formatting techniques to make the findings more reader friendly.

## Results

Overall, 2017 was a better performing year for these 12 stocks. In 2017, 11 out of the 12 stocks produced a positive return while only 2 of the 12 had positive returns in 2018. The average stock return for the 12 stocks in 2017 was 67.3% and the average return in 2018 was -8.5%. While the volumes were similar, 2017 was a much more successfull year based on the given data. I have included the analysis for both 2017 and 2018 below.
![2017 Return](/2017.png)

![2018 Return](/2018.png)

By refactoring the code, I was able to cut down on the run time. The original times to run the script was .85 seconds for both 2017 and 2018. After refactoring the code, the run times were .19 and .21 seconds for 2017 and 2018 respectively.

##Summary

###Advantages and Disadvantages of refactoring code

There are a couple of notable advantages and disadvantages to refactoring code .

Advantagages: First, it can make your code run faster. If you write a more efficient code that can take out extra steps or removes unneccessary loops, you will cut down the run time. Second, refactoring code can make your code more readable. If you need to come back to your code at a later date or if someone else needs to access your code, having a code that's easier to read makes understanding the script much easier.

Disadvantages: One disadvantage to refactoring code is that it can be very time consuming. Sometimes the time it takes to refactor your code might not make logistical sense. Second, even though your intent may be to simplify the code, that might not be the case. Sometimes refactoring code can lead you in to writing a more complex code or to a place where you get stuck and can't figure out how to complete your script.

###How these Pros and Cons apply to refacturing this VBA code

The pros of refactoring this code was that it sped up the processing time. My run time was cut down by almost 75%. Even though it's not a big deal while analyzing 12 stocks; if we want to use this code while advertising 1,000 stocks, this would save a lot of time. None of the cons I listed played in to this scenario because it didn't take long to refactor the code and I was able to complete it.
