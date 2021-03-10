---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Generate a new unique identifier and add a row to table and range

This scenario helps a user generate a unique document number with a specific format and add an entry to a range or table (two sepearate scenarios). 
The new entry or a row added will contain the newly generated unique document number and few other attributes passed to the script. 

The sample contains two part. 

* Part-1: achieves this scenario by reading and adding row to a worksheet containing plain range. 
* PArt-2: achieves this scenario by reading and adding row to a table (little bit simpler). 

## Screenshots 

### Part-1 Generate document number on a range

#### Before the new row is added 
![Range before](Range-Before.png) 
### After the new row is added 
![Range after](Range-After.png) 

### Part-2 Generate document number on a table
#### Before the new row is added 
![Table before](Table-Before.png) 
#### After the new row is added 
![Table after](Table-After.png) 

### Use this in Power Automate 

## Office Scripts source code

1. [Generate and add row to range](DocNumGenForRange.ts)
1. [Generate and add row to table](DocNumGenForTable.ts)

## Excel files used

Download the input Excel file used in this sample and try it out yourself! 

[Excel file](Document-number-generator.xlsx)
