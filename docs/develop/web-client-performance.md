---
title: 'Improve the performance of your Office Scripts'
description: 'Create faster scripts by understanding the communication between the Excel workbook and your script.'
ms.date: 06/15/2020
localization_priority: Normal
---

# Improve the performance of your Office Scripts

The purpose of Office Scripts is to automate commonly performed series of tasks to save you time. A slow script can feel like it doesn't speed up your workflow. Most of the time, your script will be perfectly fine and run as expected. However, there are a few, avoidable scenarios that can affect performance.

The most common reason for a slow script is excessive communication with the workbook. Your script runs on your local machine, while the workbook exists in the cloud. At certain times, your script synchronizes its local data with that of the workbook. This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens. Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times. In either case, the script fetches information before it acts on the data. For example, the following code will accurately log the number of rows in the used range.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary. You don't need to worry about these synchronizations for your script to run correctly. However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.

## Performance optimizations

You can apply simple techniques to help reduce the communication to the cloud. The following patterns help speed up your scripts.

- Read workbook data once instead of repeatedly in a loop.
- Remove unnecessary `console.log` statements.
- Avoid using try/catch blocks.

### Read workbook data outside of a loop

Any method that gets data from the workbook can trigger a network call. Rather than repeatedly making the same call, you should save data locally whenever possible. This is especially true when dealing with loops.

Consider a script to get the count of negative numbers in the used range of a worksheet. The script needs to iterate over every cell in the used range. To do that, it needs the range, the number of rows, and the number of columns. You should store those as local variables before starting the loop. Otherwise, each iteration of the loop will force a return to the workbook.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`. You may notice the script takes considerably longer to run when dealing with large ranges.

### Remove unnecessary `console.log` statements

Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md). However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date. Consider removing unnecessary logging statements (such as those used for testing) before sharing your script. This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.

### Avoid using try/catch blocks

We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow. Most errors can be avoided by checking objects returned from the workbook. For example, the following script checks that the table returned by the workbook exists before trying to add a row.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## Case-by-case help

As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate. If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Be sure to tag your question with "office-scripts" so experts can find it and help.

## See also

- [Scripting fundamentals for Office Scripts in Excel on the web](scripting-fundamentals.md)
- [MDN web docs: Loops and iteration](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
