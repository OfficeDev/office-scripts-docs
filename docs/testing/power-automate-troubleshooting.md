---
title: Troubleshoot Office Scripts running in Power Automate
description: Tips, platform information, and known issues with the integration between Office Scripts and Power Automate.
ms.date: 03/27/2023
ms.localizationpriority: medium
---

# Troubleshoot Office Scripts running in Power Automate

Power Automate lets you take your Office Script automation to the next level. However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.

> [!TIP]
> If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.

## Avoid relative references

Power Automate runs your script in the chosen Excel workbook on your behalf. The workbook might be closed when this happens. Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate. This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.

Some relative reference APIs throw errors in Power Automate. Others have a default behavior that implies a user's state. When designing your scripts, be sure to use absolute references for worksheets and ranges. This makes your Power Automate flow consistent, even if worksheets are rearranged.

### Script methods that fail when run in Power Automate flows

The following methods throw an error and fail when called from a script in a Power Automate flow.

| Class | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### Script methods with a default behavior in Power Automate flows

The following methods use a default behavior, in lieu of any user's current state.

| Class | Method | Power Automate behavior |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`. |

## Data refresh not supported in Power Automate

Office Scripts can't refresh data when run in Power Automate. Methods such as `PivotTable.refresh` do nothing when called in a flow. Additionally, Power Automate doesn't trigger a data refresh for formulas that use workbook links.

### Script methods that do nothing when run in Power Automate flows

The following methods do nothing in a script when called through Power Automate. They still return successfully and don't throw any errors.

| Class | Method |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## Select workbooks with the file browser control

When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow. Use the file browser to select your workbook, instead of manually typing the workbook's name.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="The Power Automate Run script action showing the Show Picker file browser option.":::

For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## Pass entire arrays as script parameters

Power Automate allows users to pass arrays to connectors as a variable or as single elements in the array. The default is to pass single elements, which builds the array in the flow. For scripts or other connectors that take entire arrays as arguments, you need to select the **Switch to input entire array** button to pass the array as one complete object. This button is in the upper-right corner of each array parameter input field.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="The button to switch to input an entire array in a control field input box.":::

## Time zone differences

Excel files don't have an inherent location or timezone. Every time a user opens the workbook, their session uses that user's local timezone for date calculations. Power Automate always uses UTC.

If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate. Power Automate allows you to convert, format, and adjust times. See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [Script parameter and return types in Power Automate](../develop/power-automate-parameters-returns.md) to learn how to provide that time information for the script.

## Script parameter fields or returned output not appearing in Power Automate

There are two reasons that the parameters or returned data of a script are not accurately reflected in the Power Automate flow builder.

- The script signature (the parameters or return value) has changed since the **Excel Business (Online)** connector was added.
- The script signature uses unsupported types. Verify your types against the [restrictions for Office Scripts parameter and return types](../develop/power-automate-parameters-returns.md).

The signature of a script is stored with the **Excel Business (Online)** connector when it is created. Remove the old connector and create a new one to get the latest parameters and return values for the **Run script** action.

## ISO strict Open XML workbooks aren't supported

The **Excel Business (Online)** connector's **Run script** action doesn't support workbooks with the [ISO strict version of the Excel Workbook file format](https://www.loc.gov/preservation/digital/formats/fdd/fdd000401.shtml). Flows with this type of workbook return a "BadGateway" error when trying to run a script. This is due to coauthoring restrictions. Please save workbooks as the standard Excel Workbook format for use with Power Automate.

## See also

- [Troubleshoot Office Scripts](troubleshooting.md)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
- [Excel Online (Business) connector reference documentation](/connectors/excelonlinebusiness/)
