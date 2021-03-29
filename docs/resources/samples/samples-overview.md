---
title: 'Office Scripts samples'
description: 'Available Office Scripts samples and scenarios.'
ms.date: 03/29/2021
localization_priority: Normal
---

# Office Scripts samples and scenarios

This section contains [Office Scripts](../../overview/excel.md) based automation solutions that help end users achieve automation of daily tasks. It contains realistic scenarios that business users face and provides detailed solutions along with step-by-step instructional video links.

For each of the projects in [Basics](#basics), [Beyond the basics](#beyond-the-basics), and [Performance](#performance), check out the source code, step-by-step [**YouTube videos**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0), and more.

In [Scenarios](#scenarios), we have included a few larger scenario samples that demonstrate real-world use cases.

We also welcome [contributions from the community](#community-contributions).

## Basics

| Project | Details |
|---------|---------|
| [Learn basics about using the Range object in Office Scripts](range-basics.md) | This article shows the basics of using Range object and its APIs. This is a foundational topic that'll be used across all other projects. |

## Beyond the basics

Check out the following end-to-end project that automates sample scenarios along with full scripts, sample Excel files used, and [videos](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0).

| Project | Details |
|---------|---------|
| [Add comments in Excel](add-excel-comments.md) | This sample shows how to add comments to a cell with at-mention of a contact. |
| [Cross reference and format an Excel file](excel-cross-reference.md) | This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate. |
| [Email a chart image](email-chart-image.md) | This sample uses Office Scripts and Power Automate actions to create a chart and send that chart as an image by email. |
| [Filter Excel table and get visible range](filter-table-get-visible-range.md) | This sample filters an Excel table and returns the visible range as a JSON object. This JSON could be provided to a Power Automate flow as part of a larger solution. |
| [Generate a new unique identifier and add a row to table and range](document-number-generator.md)  | This scenario helps a user generate a unique document number with a specific format and add an entry to a range or table. |
| [Make external API calls in Office Scripts](external-calls.md) | Office Scripts allow limited external API call support. This example script gets basic information about the user's GitHub repositories. |
| [Manage calculation mode in Excel](excel-calculation.md) | This sample shows how to use the calculation mode and calculate methods in Excel on the web using Office Scripts. |
| [Merge multiple Excel tables into a single table](copy-tables-combine.md) | This sample combines data from multiple Excel tables into a single table that includes all the rows. |
| [Move rows across tables](move-rows-across-tables.md) | This sample shows how to move rows across tables by saving filters, then processing and reapplying the filters. |
| [Output Excel data as JSON](get-table-data.md) | This solution shows how to output Excel table data as JSON to use in Power Automate. |
| [Remove hyperlinks from each cell in an Excel worksheet](remove-hyperlinks-from-cells.md) | This sample clears all of the hyperlinks from the current worksheet. |
| [Send a Teams meeting from Excel data](send-teams-invite-from-excel-data.md) | This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel. |

## Performance

| Project | Details |
|---------|---------|
| [Performance optimization when writing a large dataset](write-large-dataset.md) | Learn how to optimize performance when writing a large dataset in Office Scripts. |

## Scenarios

Office Scripts can automate parts of your daily routine. These day-to-day tasks often exist in unique ecosystems, with Excel workbooks that are set up in particular ways. These larger scenario samples demonstrate such real-world use-cases. They include both the Office Scripts and the workbooks, so you can see the scenario from end to end.

| Scenario | Details |
|---------|---------|
| [Analyze web downloads](../scenarios/analyze-web-downloads.md) | This scenario features a script that parses web traffic records to determine a user's country of origin. It showcases the skills of text parsing, using subfunctions in scripts, applying conditional formatting, and working with tables. |
| [Fetch and graph water-level data from NOAA](../scenarios/noaa-data-fetch.md) | This scenario uses an Office Script to pull data from an external source (the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/)) and graph the resulting information. It highlights the skills of using `fetch` to get data and using charts. |
| [Grade calculator](../scenarios/grade-calculator.md) | This scenario features a script that validates an instructor's record for their class's grades. It showcases the skills of error checking, cell formatting, and regular expressions. |
| [Task reminders](../scenarios/task-reminders.md) | This scenario uses an Office Script in a Power Automate flow to send reminders to coworkers to update a project's status. It highlights the skills of Power Automate integration and data transfer to and from scripts. |

## Community contributions

We welcome [contributions](../../../Contributing.md) from our Office Scripts community! Feel free to create a pull request for review.

| Project | Details |
|---------|---------|
| [Seasons greetings animation](community-seasons-greetings.md) | This script was contributed by [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) in the spirit of the holiday season! It's a fun script that shows a singing Christmas tree in Excel on the web using Office Scripts. |

## Try it out

These samples are open source. Try them out yourself. You'll need a Microsoft work or school account from work or school with a license to Microsoft 365 subscription (E3 or above). Just head over to https://office.com to sign in to your account and get started.

## Leave a comment

Feel free to leave a comment, make a suggestion, or log an issue by using the **Feedback** section at the bottom of the specific sample's documentation page.

## See also

- [Sample scripts for Office Scripts in Excel on the web](../excel-samples.md)
