# Stock Analysis using VBA

## Overview of Project
  Perform stock analysis to identify best stock for the investment. 

### Purpose
  The purpose of this project is to analyze list of green energy stock for couple of years and identify the profitable stock investment option.

## Results
### Data Analysis Results
  Below is the summarized result of green energy stock dataset.

<img width="677" alt="2017_2018_results" src="https://user-images.githubusercontent.com/83181834/117188746-40449900-ad92-11eb-8df6-c02027002916.png">

There are many green energy stock which fared well in 2017 dropped their return in 2018. Based on this analysis ENPH and RUN are the two stocks which consistently performed well in 2017 and 2018 by increasing its trading volume and the return. ENPH and RUN are good stocks to investment for better return.

### VBA Script Analysis Results

#### Code Comparison
Analysis results are achieved by two methods - Original script and refracted script.
	* Original script uses variable to store the data while looping through all records checking various conditions and printing the results for the whole dataset.
	* Refracted script follows similar approach but stores the data in arrays which loop through all the data one time to collect the same information but more efficiently.
	
  
#### Example

**Original Script snapshot**
In the below snapshot, calculation of volume, starting price , ending price by validating data for conditions and printing the data in excel all happening inside nested for loop. Since the same variables are used for each ticker we are processing in this approach but this will result in more processing time if we have million records in dataset.
	
  <img width="677" alt="Original_Script" src="https://user-images.githubusercontent.com/83181834/117189182-be08a480-ad92-11eb-80c0-c25f3f192dba.png">
  

	
**Refracted script snapshot**
In the below snapshot,  calculation of volume, starting price , ending price by validating data for conditions happening inside nested for loop and storing the data in respectable arrays. While printing, we need to access the arrays which has fewers rows than whole dataset. 

  <img width="709" alt="Refracted_code" src="https://user-images.githubusercontent.com/83181834/117189200-c2cd5880-ad92-11eb-8192-ccbefaa03cfc.png">



#### Code Execution
Capturing the execution time of code helps in analyzing the components and improve efficiency. Find below the snapshots of code execution for 2017 and 2018 dataset. Refracted code ran faster than original code because of arrays and less looping. 

<img width="677" alt="2017_results" src="https://user-images.githubusercontent.com/83181834/117189251-d4166500-ad92-11eb-9a81-db1ffd1022d2.png">
<img width="677" alt="2018_results" src="https://user-images.githubusercontent.com/83181834/117189255-d5479200-ad92-11eb-9d36-04f9012e0e6c.png">


## Summary

