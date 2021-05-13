---
title: 'Best practices in Office Scripts'
description: 'How to prevent common problems and write robust Office Scripts that can handle unexpected input or data.'
ms.date: 05/10/2021
localization_priority: Normal
---

# Best practices in Office Scripts

There are some patterns you can follow in your Office Scripts to help ensure they run successfully. These should help you avoid common pitfalls as you start automating your Excel workflow.

## Check if the object exists before using it

Scripts often rely on a certain worksheet or table being present in the workbook. However, they might get renamed or removed between script runs. By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.

The following sample code checks if the "Index" worksheet is present in the workbook. If it is, it gets a range and proceeds. If it isn't, it logs a custom error message.

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

You can also use the TypeScript `?` operator to check if the object exists before calling a method. This can make your code more streamlined if you don't need to do anything special if the object doesn't exist.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## Check everything at the beginning of the script

Make sure everything is in the workbook before working on the data. Using the previous pattern, you can check to see if all your worksheets, tables, shapes, and other objects are present and match your expectations. Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.

The following script requires two tables named "Table1" and "Table2" to be present. The script checks if they're there and ends with the `return` statement and an appropriate message if they are not.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

If the verification is happening in a separate function, you still have to end the script by issuing the `return` statement from the `main` function. Returning from the subfunction doesn't end the script.

The following script has the same behavior as the previous one. The difference is that the `main` function calls the `inputPresent` function to verify everything. `inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present. The `main` function uses that boolean to decide on continuing or ending the script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## When to use `throw` the script  

A [`throw` statement](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) indicates an unexpected error has occurred. It ends the code immediately. For the most part, you don't need to `throw` from your script. Usually, the script automatically informs the user that the script failed to run due to an issue. In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.

However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing. A `throw` statement stops the script and tells the flow to stop as well.

The following scripts shows how to use the `throw` statement in our table checking example.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## How to use try..catch to handle errors

The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) technique is a way to detect if an API call failed and handle that error in your script. It may be important to check the return value of an API to verify that it was completed successfully.

Consider the following snippet that performs a large data update on a range:

```TypeScript
range.setValues(someLargeValues);
```

If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails. The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors). You may wish to handle this condition in your code. You could either customize the error message (the first of the following two snippets) or break up the update into smaller units (the second snippet). `try...catch` lets you handle this in the script, rather than showing the default error to whoever is using your script.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> Using `try..catch` inside or around a loop slows down your script. See [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-around-loops) for more performance information.

## See also

- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Troubleshooting information for Power Automate with Office Scripts](../testing/power-automate-troubleshooting.md)
- [Platform Limits with Office Scripts](../testing/platform-limits.md)
- [Improve the performance of your Office Scripts](web-client-performance.md)
