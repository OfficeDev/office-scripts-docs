---
title: JavaScript Date samples
description: A collection of samples on how to use JavaScript Date objects with Excel.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# JavaScript `Date` samples

These samples show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.

[!INCLUDE [open-code-editor](../../includes/open-code-editor.md)]

## Write the current date and time

The following sample gets the current date and time and then writes those values to two cells in the active worksheet.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

## Read an Excel date

This sample reads a date that's stored in Excel and translates it to a JavaScript `Date` object. It uses the date's numeric serial number as input for the JavaScript `Date`. This serial number is described in the [NOW() function](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) article.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## See also

- [Use built-in JavaScript objects in Office Scripts](../../develop/javascript-objects.md)
