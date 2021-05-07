---
title: 'Count blank rows on sheets'
description: 'Learn how to use Office Scripts to detect if there are any blank rows instead of data in worksheets and then report the blank row count to be used in a Power Automate flow.'
ms.date: 05/04/2021
localization_priority: Normal
---

# Count blank rows on sheets

This project includes two scripts:

* [Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.
* [Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.

> [!NOTE]
> For our script, a blank row is any row where there's no data. The row can have formatting.

_This sheet returns count of 4 blank rows_

:::image type="content" source="../../images/blank-rows.png" alt-text="A worksheet showing data with blank rows":::

_This sheet returns count of 0 blank rows (all rows have some data)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="A worksheet showing data without blank rows":::

## Sample code: Count blank rows on a given sheet

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## Sample code: Count blank rows on all sheets

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}` );

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## Use with Power Automate

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="A Power Automate flow showing how to set up to run an Office Script":::
