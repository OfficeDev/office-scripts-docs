---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Cross reference and format Excel file

This project shows how two Excel files can be cross-refenced and formatted using Office Scripts and Power Automate. 

The project achieves this - 

1. Extracts event master (key) data from Event.xlsx using one Run script action. 
1. Passes that data to second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts. 
1. Sends the result to a reviwer via. email. 

For further details see: 

https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535


## Video link

[![Watch step by step video](v_cross_ref_tables.jpg)](https://youtu.be/dVwqBf483qo "Watch step by step video")


## Office Scripts 

Checkout the directory. 
1. [Get Event Master Data](ReturnEvents.ts)
1. [Validate Event Transactions](ValidateEventTransactions.ts)

## Excel files used

1. [Event master data](Events.xlsx)
1. [Event tranaction data](Event-Transactions.xlsx)


