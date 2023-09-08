---
title: Row and column visibility samples
description: A collection of samples that show the basics of visibility in Excel.
ms.date: 09/08/2023
ms.localizationpriority: medium
---

# Row and column visibility samples

These samples demonstrate how to show, hide, and freeze rows and columns.

## Hide columns

This script hides columns "D", "F", and "J".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  // Hide columns D, F, and J.
  sheet.getRange("D:D").setColumnHidden(true);
  sheet.getRange("F:F").setColumnHidden(true);
  sheet.getRange("J:J").setColumnHidden(true);
}
```

## Show all rows and columns

This script gets the worksheet's used range, checks if there are any hidden rows and columns, and shows them.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the entire data range.
    const range = selectedSheet.getUsedRange();

    // If the used range is empty, end the script.
    if (!range) {
      console.log(`No data on this sheet.`)
      return;
    }

    // If no columns are hidden, log message, else show columns.
    if (range.getColumnHidden() == false) {
      console.log(`No columns hidden`);
    } else {
      range.setColumnHidden(false);
    }

    // If no rows are hidden, log message, else, show rows.
    if (range.getRowHidden() == false) {
      console.log(`No rows hidden`);
    } else {
      range.setRowHidden(false);
    }
}
```

## Freeze currently selected cells

This script checks what cells are currently selected and freezes that selection, so those cells are always visible.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the current selected range.
    const selectedRange = workbook.getSelectedRange();

    // If no cells are selected, end the script. 
    if (!selectedRange) {
      console.log(`No cells in the worksheet are selected.`);
      return;
    }

    // Log the address of the selected range
    console.log(`Selected range for the worksheet: ${selectedRange.getAddress()}`);

    // Freeze the selected range.
    selectedSheet.getFreezePanes().freezeAt(selectedRange);
}
```
