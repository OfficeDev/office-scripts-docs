---
title: 'Read workbook data with Office Scripts in Excel on the web'
description: 'An Office Scripts tutorial about reading data from workbooks and evaluating that data in the script.'
ms.date: 07/20/2020
localization_priority: Priority
---

# Read workbook data with Office Scripts in Excel on the web

This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web. You'll be writing a new script that formats a bank statement and normalizes the data in that statement. As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.

Completing the tutorial expands your knowledge of Office Scripts to include reading data, instead of just writing it. This lets you base the results of the script off of the contents in the workbook, which lets you automate some of your decision making processes with the script.

> [!TIP]
> If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## Read a cell

Scripts made with the Action Recorder can only write information to the workbook. With the Code Editor, you can edit and make scripts that also read data from a workbook.

Let's make a script that reads data and acts based on what was read. We're going to work with a sample banking statement. This statement is a combined checking and credit statement. Unfortunately, they report balance changes differently. The checking statement gives income as positive credit and costs as negative debit. The credit statement does the opposite.

Over the rest of the tutorial, we will normalize this data using a script. First, let's learn how to read data from the workbook.

1. Create a new worksheet in the workbook you've used for the rest of the tutorial.
2. Copy the following data and paste it into the new worksheet, starting at cell **A1**.

    |Date |Account |Description |Debit |Credit |
    |:--|:--|:--|:--|:--|
    |10/10/2019 |Checking |Coho Vineyard |-20.05 | |
    |10/11/2019 |Credit |The Phone Company |99.95 | |
    |10/13/2019 |Credit |Coho Vineyard |154.43 | |
    |10/15/2019 |Checking |External Deposit | |1000 |
    |10/20/2019 |Credit |Coho Vineyard - Refund | |-35.45 |
    |10/25/2019 |Checking |Best For You Organics Company | -85.64 | |
    |11/01/2019 |Checking |External Deposit | |1000 |

3. Open the **Code Editor** and select **New Script**.
4. Let's clean up the formatting. This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts. Let's also fit the column width to the data.

    Replace the script contents with the following code:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. Now let's read a value from one of the number columns. Add the following code to the end of the script (before the closing `}`):

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. Run the script.
7. You should see `[Array[1]]` in the console. This is not a number because ranges are two-dimensional arrays of data. That two-dimensional range is being logged to the console directly. Luckily, the Code Editor lets you see the contents of the array.
8. When a two-dimensional array is logged to the console, it groups column values under each row. Expand the array log by pressing the blue triangle.
9. Expand the second level of the array by pressing the newly revealed blue triangle. You should now see this:

    ![The console log displaying the the output "-20.05", nested under two arrays.](../images/tutorial-4.png)

## Modify the value of a cell

Now that we can read data, let's use that data to modify the workbook. We'll make the value of the cell **D2** positive with the `Math.abs` function. The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access. More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).

1. Add the following code to the end of the script:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    Note that we're using `getValue` and `setValue`. These methods work on a single cell. When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.

2. The value of cell **D2** should now be positive.

## Modify the values of a column

Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.

1. Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. Add a loop to the end of the script that iterates through the rows in the last two columns. For each cell, the script sets the value to the current value's absolute value.

    Note that the array defining cell locations is zero-based. That means cell **A1** is `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3]);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4]);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    This portion of the script does several important tasks. First, it gets the values and row count of the used range. This lets us look at values and know when to stop. Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns. Finally, if the value in the cell is not 0, it is replaced by its absolute value. We're avoiding zeroes so we can leave the blank cells as they were.

3. Run the script.

    Your banking statement should now look like this:

    ![The banking statement as a formatted table with only positive values.](../images/tutorial-5.png)

## Next steps

Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md). You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.

The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate. Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or take the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.
