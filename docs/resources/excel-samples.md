---
title: 'Sample scripts for Office Scripts in Excel on the web'
description: 'A collection of samples to use with Office Scripts in Excel on the web.'
ms.date: 01/09/2020
localization_priority: Normal
---

# Sample scripts for Office Scripts in Excel on the web (preview)

The following samples are simple scripts for you to try on your own workbooks. To use them:

1. Open the **Automate** tab.
2. Press **Code Editor**.
3. Press **New Script** in the Code Editor's task pane.
4. Replace the entire script with the sample of your choice.
5. Press **Run** in the Code Editor's task pane.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## Read and log one cell

This sample reads the value of **A1** and prints it to the console.

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

## Create a sorted table

This sample creates a table from the current worksheet's used range, then sorts it based on the first column.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Get the current worksheet.
    let workbook = context.workbook;
    let worksheets = workbook.worksheets;
    let selectedSheet = worksheets.getActiveWorksheet();

    // Create a table with the used cells.
    let usedRange = selectedSheet.getUsedRange();
    let newTable = selectedSheet.tables.add(usedRange, true);

    // Sort the table using the first column.
    newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## Delete resolved comments

This sample deletes all resolved comments from the current worksheet.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Get the current worksheet.
    let workbook = context.workbook;
    let worksheets = workbook.worksheets;
    let selectedSheet = worksheets.getActiveWorksheet();

    // Get the comments on this worksheet.
    let comments = selectedSheet.comments;
    comments.load("items/resolved");
    await context.sync();

    // Delete the resolved comments.
    comments.items.forEach((comment) => {
        if (comment.resolved) {
            comment.delete();
        }
    });
}
```

## Apply conditional formatting
