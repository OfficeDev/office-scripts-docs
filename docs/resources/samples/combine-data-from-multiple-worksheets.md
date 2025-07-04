---
title: Combine data from multiple worksheets
description: Learn how to use Office Scripts to merge data from other worksheets into a single worksheet.
ms.date: 07/03/2025
ms.localizationpriority: medium
---

# Combine data from multiple worksheets

Use Office Scripts to combine data from multiple worksheets into a single worksheet, creating a unified report of your data.

> [!IMPORTANT]
> This sample only copies the values from the other worksheets. It does not preserve formatting, charts, tables, or other objects.

## Solution

1. Create a new Excel file in your OneDrive.
1. Add data to multiple worksheets.
1. Create the script from this sample.
1. Run the script.

## Sample code: Combine data

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet, and then combines them on a new worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
    // Create an object to return the data from each worksheet.
    let worksheetInformation: WorksheetData[] = [];

    // Get the data from every worksheet, one at a time.
    workbook.getWorksheets().forEach((sheet) => {
        let values = sheet.getUsedRange()?.getValues();
        worksheetInformation.push({
            name: sheet.getName(),
            data: values as string[][]
        });
    });
  
    // Create the new worksheet.
    let sheet = workbook.addWorksheet("Combined Data");

    // Add data from each worksheet to new worksheet.
    worksheetInformation.forEach((value) => {
        // If there was any data in the worksheet, add it to a new range.
        if (value.data) {
            let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
            range.insert(ExcelScript.InsertShiftDirection.down);
            range.setValues(value.data);
        }
    });    
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
    name: string;
    data: string[][];
}
```

## Next steps

Learn how to [save your report as a PDF and email it](save-and-email-as-pdf.md) to yourself or your team.
