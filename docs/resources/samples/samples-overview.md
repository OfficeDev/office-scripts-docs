---
title: Office Scripts samples
description: Available Office Scripts samples and scenarios.
ms.date: 09/20/2023
ms.localizationpriority: medium
---

# Office Scripts samples and scenarios

This section contains [Office Scripts](../../overview/excel.md) based solutions that help end users achieve automation of daily tasks. It contains realistic scenarios that users face and provides detailed solutions.

- [Basics](#basics) showcase common examples that make up larger scripts. Most don't need a specific workbook or dataset and can be run in the workbook of your choice.
- [Beyond the basics](#beyond-the-basics) are samples that are more involved or solve a particular problem. Some use Power Automate with an Office Script as a integral part of the flow.
- [Scenarios](#scenarios) are a few larger samples that demonstrate real-world use cases.
- [Contributions from the community](#community-contributions-and-fun-samples) are samples from members of the Office Scripts community, often light-hearted in nature.

> [!IMPORTANT]
> Be sure you meet the prerequisites for Office Scripts before trying the samples. The requirements for your Microsoft 365 subscription and account are found under the [Office Scripts for Excel overview "Requirements" section](../../overview/excel.md#requirements).

## Basics

| Project | Details |
|---------|---------|
| [Range samples](range-samples.md) | These samples show how to work the `Range` object, which is central to most scripts. |
| [Table samples](table-samples.md) | A collection of samples that show common interactions with Excel tables.
| [Add comments in Excel](add-excel-comments.md) | This sample adds comments to a cell including @mentioning a colleague. |
| [Add images to a workbook](add-image-to-workbook.md) | This sample adds an image to a workbook and copies an image across sheets.|
| [Copy multiple Excel tables into a single table](copy-tables-combine.md) | This sample combines data from multiple Excel tables into a single table that includes all the rows. |
| [Data validation: dropdown lists, prompts, and warning pop-ups](data-validation-samples.md) | These samples show how to use data validation to mandate specific conditions for cell data and how the user is alerted to these rules. |
| [Create a workbook table of contents](table-of-contents.md) | This sample creates a table of contents with links to each worksheet. |
| [JavaScript `Date` samples](javascript-dates.md) | A collection of samples that show how to translate between JavaScript and Excel date formats. |
| [Record day-to-day changes in Excel and report them with a Power Automate flow](report-day-to-day-changes.md) | This sample uses a scheduled Power Automate flow to record daily readings and report the changes. |
| [Row and column visibility samples](row-and-column-visibility.md) | A collection of samples that demonstrate how to show, hide, and freeze rows and columns. |
| [Set conditional formatting for cross-column comparisons](conditional-formatting-parameters.md) | This sample applies formatting based on values in adjacent columns. It also gets user input through script parameters. |

## Beyond the basics

Check out the following end-to-end project that automates sample scenarios along with full scripts, sample Excel files used, and [videos (hosted on YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Details |
|---------|---------|
| [Combine worksheets into a single workbook](combine-worksheets-into-single-workbook.md) | This sample uses Office Scripts and Power Automate to pull data from other workbooks into a single workbook. |
| [Convert CSV files to Excel workbooks](convert-csv.md) | This sample uses Office Scripts and Power Automate to create .xlsx files from .csv files. |
| [Cross-reference workbooks](excel-cross-reference.md) | This sample uses Office Scripts and Power Automate to cross-reference and validate information in different workbooks. |
| [Count blank rows in a specific sheet or in all sheets](count-blank-rows.md) | This sample detects if there are any blank rows in sheets where you anticipate data to be present and then report the blank row count for usage in a Power Automate flow. |
| [Email chart and table images](email-images-chart-table.md) | This sample uses Office Scripts and Power Automate actions to create a chart and send that chart as an image by email. |
| [Manage calculation mode in Excel](excel-calculation.md) | This sample shows how to use the calculation mode and calculate methods in Excel using Office Scripts. |
| [Move rows across tables](move-rows-across-tables.md) | This sample shows how to move rows across tables by saving filters, then processing and reapplying the filters. |
| [Output Excel data as JSON](get-table-data.md) | This solution shows how to output Excel table data as JSON to use in Power Automate. |
| [Remove hyperlinks from each cell in an Excel worksheet](remove-hyperlinks-from-cells.md) | This sample clears all of the hyperlinks from the current worksheet. |
| [Run a script on all Excel files in a folder](automate-tasks-on-all-excel-files-in-folder.md) | This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business (can also be used for a SharePoint folder). It performs calculations on the Excel files, adds formatting, and inserts a comment that @mentions a colleague. |
| [Use external fetch calls](external-fetch-calls.md) | This sample uses `fetch` to get information from GitHub for the script. |
| [Write a large dataset](write-large-dataset.md) | This sample shows how to send a large range as smaller subranges. |

## Scenarios

Office Scripts can automate parts of your daily routine. These day-to-day tasks often exist in unique ecosystems, with Excel workbooks that are set up in particular ways. These larger scenario samples demonstrate such real-world use-cases. They include both the Office Scripts and the workbooks, so you can see the scenario from end to end.

| Scenario | Details |
|---------|---------|
| [Analyze web downloads](../scenarios/analyze-web-downloads.md) | This scenario features a script that parses web traffic records to determine a user's country of origin. It showcases the skills of text parsing, using subfunctions in scripts, applying conditional formatting, and working with tables. |
| [Fetch and graph water-level data from NOAA](../scenarios/noaa-data-fetch.md) | This scenario uses an Office Script to pull data from an external source (the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/)) and graph the resulting information. It highlights the skills of using `fetch` to get data and using charts. |
| [Grade calculator](../scenarios/grade-calculator.md) | This scenario features a script that validates an instructor's record for their class's grades. It showcases the skills of error checking, cell formatting, and regular expressions. |
| [Schedule interviews in Teams](../scenarios/schedule-interviews-in-teams.md) | This scenario shows how to use an Excel spreadsheet to manage interview meeting times and make a flow to schedules meetings in Teams. |
| [Task reminders](../scenarios/task-reminders.md) | This scenario uses an Office Script in a Power Automate flow to send reminders to coworkers to update a project's status. It highlights the skills of Power Automate integration and data transfer to and from scripts. |

## Community contributions and fun samples

We welcome [contributions](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) from our Office Scripts community! Feel free to create a pull request for review.

| Project | Details |
|---------|---------|
| [Game of Life](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | The "Ready Player Zero" blog by Yutao Huang on the Excel Tech Community includes a script to model John Conway's [*The Game of Life*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life). |
| [Punch clock button](../scenarios/punch-clock.md) | This script was contributed by [Brian Gonzalez](https://github.com/b-gonzalez). The scenario features a script and a script button that records the current time. |
| [Seasons greetings animation](community-seasons-greetings.md) | This script was contributed by [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) in the spirit of the holiday season! It's a fun script that shows a singing Christmas tree in Excel using Office Scripts. |

## Leave a comment

Feel free to leave a comment, make a suggestion, or log an issue by using the **Feedback** section at the bottom of the specific sample's documentation page.
