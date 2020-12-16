---
title: 'Troubleshooting information for Power Automate with Office Scripts'
description: 'Tips, platform information, and known issues with the integration between Office Scripts and Power Automate.'
ms.date: 12/16/2020
localization_priority: Normal
---

# Troubleshooting information for Power Automate with Office Scripts

Power Automate lets you take your Office Script automation to the next level. However, because Power Automate will be running scripts on your behalf in independent Excel sessions, there are a few important things to note.

> [!TIP]
> If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.

## Avoid using relative references

Power Automate runs your script in the chosen Excel workbook on your behalf. The workbook might be closed when this happens. Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate. When designing your scripts, be sure to use absolute references for worksheets and ranges.

The following methods will throw an error and fail when called from a script in a Power Automate flow.

| Class | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## Select workbooks with the file browser control

When building the **Run script** step of a Power Automate flow, you'll need to select which workbook is part of the flow. Use the file browser to select your workbook, instead of manually typing the workbook's name.

![The file browser option when creating a "Run script" action in Power Automate](../images/power-automate-file-browser.png)

For more context of the Power Automate limitation and a discussion of potential workarounds for dynamic selections of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)

## Time zone differences

Excel files don't have an inherent location or timezone. Every time a user opens the workbook, their session uses that user's local timezone when needed for dates. Power Automate always uses UTC.

If your scripts uses dates or time, there may be behavioral differences when the script is tested locally and when it is ran through Power Automate. Power Automate allows you to convert, format, and adjust times. See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) to learn how to provide that time information for the script.

## See also

- [Troubleshooting Office Scripts](troubleshooting.md)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
- [Excel Online (Business) connector reference documentation](/connectors/excelonlinebusiness/)
