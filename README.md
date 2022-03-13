# About
The script uses Python to simplify complicated calculation.
Package use: Pandas, openpyxl, numpy
# What the script does - Summary
* Import data from Excel sheet to pandas dataframe
* Create many lists of calendar dates that fit multiple criteria. 
* Perform calculation
* Export result to excel

# Context
My team was managing relationships between a list of organisations and their suppliers. My task was to forecast the income of each suppliers each month, based on their estimated daily income.
## About the data
* It's a many-to-many relationship model where each suppliers provide to multiple customers and vice versa.
* New suppliers start and end contracts through-out the year.
* The customer organistations only operate on certain periods of the year, excluding weekend and holiday.
* The customers can choose how many days in a week to receive service.
## Problem
The calculation has to find which days of each month the organisation will receive the service from a particular supplier. This need to takes into account the date that the contracts start and end the dates that the orgnisations close. The Excel formular become too long and complicated to track and correct. It's also hard to apply on the new datasets in the future.
# What the script does
* Import data
* Create a list of service delivery dates in each month for each unique pair of customer and supplier. The list need to meet the below criteria:
  * On or after the contract start date
  * On or before the contract end date
  * The customer organisations are operating
 * Count the number of dates in each list
 * Calculate the income for the month
 * Sum the income by supplier
 * Export the result
