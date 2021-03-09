# Module2_VBA-Challenge
An analysis of stock performance using VBA

## **Purpose of this Analysis**

###
Steve wants to help his parents understand which stocks to invest in, since DQ proved not to be a lucrative investment. Steve has provided a data set with information on thousands of stock from the last few years, and wants to be able to analyze the entire dataset at the click of a button, and perform the analysis quickly. 

## **Comparison between 2017 and 2018 Stock Performance**
<img width="535" alt="Original_2017_stock-performance" src="https://user-images.githubusercontent.com/69849998/110401621-e2245f80-8047-11eb-99aa-2580a32dbcc7.png">

<img width="539" alt="Original_2018_stock-performance" src="https://user-images.githubusercontent.com/69849998/110401627-e6e91380-8047-11eb-8fd7-5ec645fdd7bf.png">

* Overall, a comparison of the green vs red colours in the "Return" column of each of the 2017 and 2018 charts demonstrates that the % return on stocks generally was higher (green) in 2017 than 2018. 

  More specifically, in 2018, ENPH & RUN were the only two tickers that had returns above 0%, at 81.92% and 83.95% respectively. In 2017, all 12 stocks in the data performed above 0% except TERP, which returned -7.21%. The highest performing stocks were DQ AND SEDG, which had returns of 199.45% and 184.47% respectively 


<img width="251" alt="Original_2017_Execution-time" src="https://user-images.githubusercontent.com/69849998/110402097-c66d8900-8048-11eb-8516-4b27b7a0acde.png">

![Refactored_2017_Execution-time](https://user-images.githubusercontent.com/69849998/110402111-cbcad380-8048-11eb-8b98-911288e177c5.png)
* The run time for both the 2017 and 2018 analyses in the original script was 0.5 seconds, and with the refactored code, it was about 0.09 seconds to analyze the 2017 and 2018 data with the refactored code. Since the run time was lower for the refactored code, compared to the original code, this likely means that the refactored code is more efficient. 


## **Summary**
### 1. What are the advantages or disadvantages of refactoring code?

Refactoring code is advantageous because it can reduce the amount of time spent writing code, and can also add more functionality and efficiency to the code, like the stock performance analysis. The disadvantage is that it is easier to make overlook updating certain aspects of the code for the new purpose. 


### 2. How do these pros and cons apply to refactoring the original VBA script?

Similarly, for refactoring the original VBA script, it was faster to write the code, and the code is more efficient, because it has a faster run time, and has the potential to analyze more data. On the other hand, I ran in to a few minor time-consuming errors in refactoring the code. 
