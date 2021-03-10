---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Send Teams meeting from Excel data 

This project shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send Teams meeting invite and update back Excel. 

## Video link

[![Watch step by step video](v_teams_invite.jpg)](https://youtu.be/HyBdx52NOE8 "Watch step by step video")

## Scenario: 

Example scenario:

* A HR recruiter manages the interview schedule of candidates in a Excel file.
* Recruiter needs to send Teams meeting invite out to candidate and interviewers to eligible candidates. The business rules are to select: (a) invites to only those for whom the invite is already not sent as recorded in the Excel column (b) interview date is in the future (no past dates).
* Recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records. 


## Solution 

The solution has 3 parts: 

1. Office Script to extract data from Excel table based on conditions and returns an array of objects as JSON data. 
1. The data is then sent to Teams **Create a Teams meeting** action to send invites. Send one Teams meeting per instance in the JSON array. 
1. Send the same JSON data to another Office Script to update the status of the invitation. 

## Input Excel file

This is the [input Excel file](HR-Schedule.xlsx) being used. 

## Office Scripts

The solution uses two Office Scripts. 

1. [Select filtered rows from table as JSON](SelectFilteredRowsFromTableAsJSON.ts)
1. [Mark as invited](MarkAsInvited.ts)


