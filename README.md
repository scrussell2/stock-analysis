# An Analysis Review of Green Stocks VBA Refactored Code

### Green Stocks VBA Refactor Analysis Project Overview

This README analysis of refactoring VBA (*Visual Basic for Applications*) code in the attached Green Stocks Excel document for the year of 2018 and determining if there were any gained efficiencies in making the VBA script run faster.  In this analysis, there will be snippets of code for comparisons to showcase what was utilized during the refactoring along with images of the timer before and post-refactoring.

### Purpose and Background

#### Purpose

This code refactoring aims to ensure that the client who requested the Green Stocks Excel document to showcase certain stock information can be analyzed with the given datasets efficiently without taking a lot of time to process.  Additionally, a request has been made to analyze not only one year's worth of data, but multiple years, and along with buttons inside the Excel document for ease of use.

#### Background

The client for which this refactoring has been completed for is a recent graduate, Steve, who has attained his finance degree.  His parents are requesting Steve about investment opportunities in green energy companies for diversification of funds.  The Green Stocks Excel document was given by Steve for help in analyzing the data, and in good faith to be used as a repeatable product in future years as more data is collected.  VBA was chosen as the programming language to accurately perform calculations for the analysis of data by interacting with Excel.  This also allows additional datasets to be included when available from the client without further knowledge of using Excel's internal formulas.  This is accomplished by having scripts able to run in the background with the push of a button and input by the client for ease of use.

### Results

To answer directly, there was improvement between the original and refactored code by 145% difference between the two.  The original code ran at 0.7382813 seconds, and the refactored at 0.1171875 seconds which can be seen in figure 1:

**Figure 1**: "Excel Formula YEAR()"
	
	=YEAR(@S:S)

The majority of these campaigns did come from the US at 900, with the next largest from Great Britain at 353.  Regardless, they both averaged relatively similarly and had the biggest impact on the given data displayed in Image 2.

**Image 2**: "Theater Ouctomes Based on Launch Date"

### Summary

#### Advantages and Disadvantages of Refactoring Code

#### Advantages and Disadvantages of the Original and Refactored VBA Script
