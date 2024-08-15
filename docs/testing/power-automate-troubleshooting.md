---
title: Troubleshoot Office Scripts running in Power Automate
description: Tips, platform information, and known issues with the integration between Office Scripts and Power Automate.
ms.topic: troubleshooting-general
ms.date: 08/13/2024
ms.localizationpriority: medium
---

# Troubleshoot Office Scripts running in Power Automate

Power Automate runs scripts on your behalf in independent Excel sessions. This causes some behavioral changes that may create issues with certain scripts or scenarios. There are also limitations and behaviors from the Power Automate platform script writers should know. Be sure to read the articles [Troubleshoot Office Scripts](troubleshooting.md) and [Platform limits and requirements with Office Scripts](platform-limits.md), as much of that information also applies to scripts in flows.

> [!TIP]
> If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.

[!INCLUDE [Power Automate needs a business license](../includes/power-automate-needs-business.md)]

## Avoid relative references

Power Automate runs your script in the chosen Excel workbook on your behalf. The workbook might be closed when this happens. Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate. This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.

Some relative reference APIs throw errors in Power Automate. Others have a default behavior that implies a user's state. When designing your scripts, be sure to use absolute references for worksheets and ranges. This makes your Power Automate flow consistent, even if worksheets are rearranged.

### Script methods that fail in Power Automate flows

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

## Refresh not fully supported in Power Automate

Office Scripts can't refresh most data when run in Power Automate. Most refresh methods, such as `PivotTable.refresh`, do nothing when called in a flow. `Workbook.refreshAllDataConnections` only refreshes when PowerBI is the source. Additionally, Power Automate doesn't trigger a data refresh for formulas that use workbook links.

### Script methods that do nothing in Power Automate flows

The following methods do nothing in a script when called through Power Automate. They still return successfully and don't throw any errors.

| Class | Method |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

### Script methods with a different behavior in Power Automate

The following methods act differently in Power Automate flows than they do when run through Excel.

| Class | Method | Power Automate behavior |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` | Only refreshes PowerBI sources. For other sources, the method returns successfully but does nothing. |

## Select workbooks with the file browser control

When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow. Use the file browser to select your workbook, instead of manually typing the workbook's name.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="The Power Automate Run script action showing the Show Picker file browser option.":::

For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## Pass entire arrays as script parameters

Power Automate allows users to pass arrays to connectors as a variable or as single elements in the array. The default is to pass single elements, which builds the array in the flow. For scripts or other connectors that take entire arrays as arguments, you need to select the **Switch to input entire array** button to pass the array as one complete object. This button is in the upper-right corner of each array parameter input field.

:::image type="content" source="../images/combine-worksheets-flow-4.png" alt-text="The button to switch to input an entire array in a control field input box.":::

## Time zone differences

Excel files don't have an inherent location or timezone. Every time a user opens the workbook, their session uses that user's local timezone for date calculations. Power Automate always uses UTC.

If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate. Power Automate allows you to convert, format, and adjust times. See [Working with Dates and Times inside of your flows](https://make.powerautomate.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [Pass data to and from scripts in Power Automate](../develop/power-automate-parameters-returns.md) to learn how to provide that time information for the script.

## Script parameter fields or returned output not appearing in Power Automate

There are two reasons that the parameters or returned data of a script are not accurately reflected in the Power Automate flow builder.

- The script signature (the parameters or return value) has changed since the **Excel Business (Online)** connector was added.
- The script signature uses unsupported types. Verify your types against the [restrictions for Office Scripts parameter and return types](../develop/power-automate-parameters-returns.md).

The signature of a script is stored with the **Excel Business (Online)** connector when it is created. Remove the old connector and create a new one to get the latest parameters and return values for the **Run script** action.

## Some web APIs not available with Power Automate flows

Some web APIs, such as `TextEncoder` and `Crypto`, may not be available when running Office Scripts in Power Automate flows. See [MDN Web APIs](https://developer.mozilla.org/docs/Web/API) for a full list of web APIs.

Power Automate returns the error `*API* is not defined`, where `*API*` specifies a library such as `TextEncoder`, when running a script that uses an unsupported API.

## See also

- [Troubleshoot Office Scripts](troubleshooting.md)
- [Run Office Scripts with Power Automate](../develop/power-automate-integration.md)
- [Excel Online (Business) connector reference documentation](/connectors/excelonlinebusiness/)
