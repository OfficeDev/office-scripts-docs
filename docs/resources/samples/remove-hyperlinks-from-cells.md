---
title: Remove hyperlinks from each cell in an Excel worksheet
description: Learn how to use Office Scripts to remove hyperlinks from each cell in an Excel worksheet.
ms.date: 06/29/2021
ms.localizationpriority: medium
---

# Remove hyperlinks from each cell in an Excel worksheet

 This sample clears all of the hyperlinks from the current worksheet. It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is. Also logs the time it takes to complete traversal.

> [!NOTE]
> This only works if the cell count is < 10k.

## Sample Excel file

> [!div class="nextstepaction"]
> [Download the sample workbook](remove-hyperlinks.xlsx)

Add the following script to try the sample yourself!

## Sample code: Remove hyperlinks

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## Training video: Remove hyperlinks from each cell in an Excel worksheet

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).
