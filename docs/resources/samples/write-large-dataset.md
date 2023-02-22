---
title: Write a large dataset
description: Learn how to split a large dataset into smaller write operations in Office Scripts.
ms.date: 02/22/2023
ms.localizationpriority: medium
---

# Write a large dataset

The `Range.setValues()` API puts data in a range. This API has limitations depending on various factors, such as data size and network settings. This means that if you attempt to write a massive amount of information to a workbook as a single operation, you'll need to write the data in smaller batches in order to reliably update a [large range](../../testing/platform-limits.md).

The first part of the sample shows how to write a large dataset only within Excel. The second part expands the example to be part ofa Power Automate flow. This is necessary if your script takes longer to run that the [Power Automate action timeout](../../testing/platform-limits.md#power-automate).

For performance basics in Office Scripts, please read [Improve the performance of your Office Scripts](../../develop/web-client-performance.md).

## Sample 1: Write a large dataset in batches

This script writes rows of a range in smaller parts. It selects 1000 cells to write at a time. Run the script on a blank worksheet to see the update batches in action. The console output gives further insight into what's happening.

> [!NOTE]
> You can change the number of total rows being written by changing the value of `SAMPLE_ROWS`. You can change the number of cells to write as a single action by changing the value of `CELLS_IN_BATCH`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const SAMPLE_ROWS = 100000;
  const CELLS_IN_BATCH = 10000;

  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  console.log(`Generating data...`)
  let data: (string | number | boolean)[][] = [];
  // Generate six columns of random data per row. 
  for (let i = 0; i < SAMPLE_ROWS; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInBatches(sheet.getRange("B2"), data, CELLS_IN_BATCH);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInBatches(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInBatch: number
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per batch setting: ${cellsInBatch}`);

  // Determine the total number of cells to write.
  const totalCells = values.length * values[0].length;
  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInBatch) {
    console.log(`No need to batch -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  // Determine how many rows to write at once.
  const rowsPerBatch = Math.floor(cellsInBatch / values[0].length);
  console.log("Rows per batch: " + rowsPerBatch);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let batchCount = 0;

  // Write each batch of rows.
  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerBatch) {
      batchCount++;
      console.log(`Calling update next batch function. Batch#: ${batchCount}`);
      updateNextBatch(startCell, values, rowsPerBatch, totalRowsUpdated);

      // Write a completion percentage to help the user understand the progress.
      rowCount = 0;
      totalRowsUpdated += rowsPerBatch;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);
    }
  }
  
  console.log(`Updating remaining rows -- last batch: ${rowCount}`)
  if (rowCount > 0) {
    updateNextBatch(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInBatch).toFixed(8)} seconds per ${cellsInBatch} cells-batch.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */
function updateNextBatch(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerBatch: number,
  totalRowsUpdated: number
) {
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerBatch - 1, data[0].length - 1);
  console.log(`Updating batch at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerBatch);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the batch range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

### Training video: Write a large dataset

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/BP9Kp0Ltj7U).

## Sample 2: Write data in batches from a Power Automate flow

For this sample, you'll need to complete the following steps.

1. Create a workbook in OneDrive named **SampleData.xlsx**.
1. Create a second workbook in OneDrive named **TargetWorkbook.xlsx**.


### Sample code: Read part of a workbook

```TypeScript
function main(workbook: ExcelScript.Workbook, startRow: number, batchSize: number) : string[][] {
    // This sample only reads the first worksheet in the workbook.
    const sheet = workbook.getWorksheets()[0];

    // Get the boundaries of the range.
    // Note that we're assuming usedRange is too big to read or write as a single Range.
    const usedRange = sheet.getUsedRange();
    const columnCount = usedRange.getColumnCount();
    const rowCount = usedRange.getRowCount();

    // If we're starting past the last row, exit the script.
    if (startRow >= rowCount) {
        return [[]];
    }

    // Get the next batch or the rest of the rows, whichever is smaller.
    const rowCountToRead = Math.min(batchSize, (rowCount - (startRow + batchSize)));
    const rangeToRead = sheet.getRangeByIndexes(startRow, 0, rowCountToRead, columnCount);
    return rangeToRead.getValues() as string[][];
}
```

### Sample code: Write part of a workbook

```TypeScript
function main(workbook: ExcelScript.Workbook, data: string[][], currentRow: number, batchSize: number): boolean {
  // Get the first worksheet.
  const sheet = workbook.getWorksheets()[0];

  // Set the given data.
  if (data.length > 0) {
    sheet.getRangeByIndexes(currentRow, 0, data.length, data[0].length).setValues(data);
  }

  // If we wrote less data than the batch size, signal the end of the flow.
  return batchSize > data.length;
}
```

### Power Automate flow: Read and write data in a loop

