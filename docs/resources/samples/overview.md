---
title: 'Differences between Office Scripts and Office Add-ins'
description: 'The behavior and API differences between Office Scripts and Office Add-ins.'
ms.date: 02/22/2021
localization_priority: Normal
---

# Office Scripts projects

This repository contains [Office Scripts](https://docs.microsoft.com/office/dev/scripts/overview/excel) based automation solutions that helps end users achieve automation of daily tasks. It contains realistic scenarios that business users face and provides detailed solution along with step-by-step instructional video links.

For each of the following projects, check out the folders for description, source code, step-by-step [**YouTube videos**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) for select projects.

## Basics

| Project | Details |
|---------|---------|
| [Basics of Office Script and getting started](Getting%20Started)    | Learn about the Office Script language, basics of the object model, performance tips, error handling, etc.   |
| [Learn about the Excel Range object](Range%20Basics)    | This article shows the basics of using Range object and its APIs. This is a foundational topic that'll be used across all other projects. Start here...   |

[Official documentation site](https://docs.microsoft.com/office/dev/scripts/)

## Beyond Basics

Check out the following end to end project that automates sample scenarios along with full scripts, sample Excel files used, and [videos](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Details |
|---------|---------|
| [Excel Recalc, create chart, extract chart, table image, email](Chart%20and%20Table%20Images)    | This project shows how to use Office Scripts and Power Automate actions to re-calculate, create chart and extract images as base-64 string to be used later to send Email.        |
| [Move select rows across tables using filter and value lookup (2 scripts)](Move%20Rows%20Across%20Tables)    | This project moves certain rows that meet a criteria from one table (source) into a another table (target) on another worksheet using (1) plain values based selection and (2) using column filter.     |
|[Copy Excel tables into new master table](Copy%20Tables%20to%20Master%20Table)|This project shows how to combine data that resides in Excel tables into a new master table with all rows included. It creates a new worksheet and combines all rows into a single table.|
| [Filter Table and get Visible range](Filter%20Table%20Get%20Visible%20Range%20as%20Object%20Array)    | This project shows how to use Office Scripts to filter Excel table to get list of rows visible as array of objects.        |
| [Document number generator](Document%20Number%20Generator)    | This scenario helps a user generate a unique document number with a specific format and add an entry to a range or table  |
| [Add comments to Excel cell with at-mention of a contact](Add%20Excel%20Comments)    | This project shows how to use Office Scripts to add comments to Excel cell with at-mention of a contact.|
| [Learn how to manage calculation mode and calculate in Excel for web](Excel%20Calculation)    | This short project shows how to use Office Scripts to manage calculation mode in Excel for web and calculate in Excel files whose mode is set to be in manual mode.|
| [Cross reference Excel files and format Excel cells](Event%20Cross%20Reference)    | This project shows how two Excel files can be cross-refenced and formatted using Office Scripts and Power Automate. |
| [Send Teams meeting invite and update Excel](Excel%20and%20Teams%20Invite)    |  This project shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send Teams meeting invite and update back Excel. |
| [Remove hyperlink from Excel cells](Remove%20Hyperlinks%20from%20Excel%20Cells)    | This sample clears all of the hyperlink from the current worksheet. It traverses through the worksheet and if there is any hyperlink associated with the cell, it cleares the hyperlink and retains the cell value as is. Logs the time it takes to complete traversal. |
| [Output Excel table data as array of objects (as JSON) with hyperlink data for usage in Power Automate](Return%20Table%20Data%20as%20Array%20of%20Objects)    | Output Excel table data as array of objects (as JSON) for usage in Power Automate. Contains two samples - one that extracts basic text values and a second one that extracts hyperlink value instead of the text value for one of the columns. |
| [External API Calls using Office Scripts](API%20Calls)| Office Scripts allows limited external API call support. Learn about the limitations and basics of how to make an external call. |
| [Simple technique to use Excel macro files (xlsm) in Power Automate Run Script action ](Excel%20and%20Power%20Automate/Using%20Excel%20Macro%20Files.MD)    | Get around the limitation to use xlsm files in Excel online's Run Script action. |



## Performance 

| Project | Details |
|---------|---------|
| [Performance related topics](Performance)    | Basic to advanced performance related topics (..in progress..)   |

## Community contributions 

| Project | Details |
|---------|---------|
| [Seasons greetings animation project](Community%20Projects/Seasons%20Greetings)    | This is a script contributed by [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) that creates Holidays theme animation in Excel using Office Scripts.   |

## Try it out yourself

The project is open sourced. Try it out yourself. You'll need Microsoft (Office) 365 account from work or school with E3 or above license. Just head over to https://office.com to sign-in to your account and get started.

## Contact

I work for the Office Scripts product team. I'm eager to get feedback on the features I work on. I'm also passionate about automation, web based solutions, TypeScript/JavaScript technologies and hence my motiviation to add these content both as a way to test the product I work on and also contribute samples that others can use. I welcome any feedback you can send, positive or negative, on the product/code/content. Feel free to log your comments in issues. 

* [LinkedIn](https://www.linkedin.com/in/rsudhi/)
* [Twitter](https://twitter.com/rsudhi)


## Leave a comment
Feel free to leave a comment or make a suggestion or log an issue by [opening a new issue](https://github.com/sumurthy/officescripts-projects/issues).
