---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Remove hyperlinks from each cell in a Excel worksheet 

 This sample clears all of the hyperlink from the current worksheet. It traverses through the worksheet and if there is any hyperlink associated with the cell, it cleares the hyperlink and retains the cell value as is. 
 Logs the time it takes to complete traversal.
 
 *Note that this only works if the cell count is < 10k*
 
## Video 

[![Watch step by step video](v_hyperlinks.jpg)](https://youtu.be/v20fdinxpHU "Watch step by step video")

## Scenario

If your Excel sheet contains many hyperlinks that needs to be removed, automating that action will save time. 

Checkout the script and video for more details. 

## Input Excel file

[Excel file with Hyperlinks](Remove-Hyperlinks.xlsx)

## Office Scripts

1. [Remove hyperlinks from Excel for web](RemoveHyperlink.ts)
