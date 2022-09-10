# Stock-analysis

## Overview
The purpose of this project was to refactor lines of VBA to return information based on the date of stocks from 2017 and 2018. 
The idea was to refactor the code to improve effiency when comparing against larger data sets. Based on 12 tickers, we were able to find the volume and return of each stock from both 2017 and 2018, allowing us to easily compare the two years.

## Results

To get our results, I had to create a For loop to go over all rows on both the 2017 and 2018 worksheets. Next I had to create If Then statements to compare rows of data, the result is shown below.

<img width="668" alt="Screen Shot 2022-09-10 at 12 00 46 PM" src="https://user-images.githubusercontent.com/109381647/189497944-7d8e99d4-230c-40cd-8a09-de0f2ae4a5f6.png">


<img width="258" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/109381647/189497859-28d03b48-af0a-4d1c-b7df-5d48efbf5c54.png">
<img width="257" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/109381647/189497866-8ca9933b-ec6a-4ccb-9abc-f02900757f11.png">

## Summary

When it comes to refactoring code, there are advantages and disadvantages. One advantage to redactoring is you are already provided with a base to go off of. You take code that may already work, and you modify it to be more effiecient. A disadvantage to refactoring is trying to determine what needs to be changed. If you try to change a line, it could potentially create more bugs and errors you will have to comb through.

Refactoring the VBA script for this project also had it's advantages and disadvantages. One positive to the original VBA script is it's simplicity, however the disadvantage is it would not work well for larger data sets. That is why we refactored the code, to read larger sets of data in a faster time.  
