---
title: "Tutorial: Create and format an Excel table"
description: A tutorial about the basics of Office Scripts, including recording scripts with the Action Recorder and writing data to a workbook.
ms.date: 09/20/2024
ms.localizationpriority: high
---

# Tutorial: Create and format an Excel table

This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel. You'll record a script that applies some formatting to a sales record worksheet. You'll then edit the recorded script to apply more formatting, create a table, and sort that table. This record-then-edit pattern is an important tool to see what your Excel actions look like as code.

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.

## Add data and record a basic script

First, you'll need some data and a small starting script.

1. Create a new Excel workbook.
1. Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.

    |Fruit |2018 |2019 |
    |:---|:---|:---|
    |Oranges |1000 |1200 |
    |Lemons |800 |900 |
    |Limes |600 |500 |
    |Grapefruits |900 |700 |

1. Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by selecting the drop-down arrow. If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).
1. Select the **Record Actions** button.
1. Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.
1. Stop the recording by selecting the **Stop** button.

    Your worksheet should look like this (don't worry if the color is different):

    :::image type="content" source="../images/tutorial-1.png" alt-text="A worksheet showing fruit sales data row with the row containing 'Oranges' highlighted in the color orange.":::

## Edit an existing script

The previous script colored the "Oranges" row to be orange. Add a yellow row for the "Lemons".

1. From the now-open **Details** pane, select the **Edit** button.
1. You should see something similar to this code:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    This code gets the current worksheet from the workbook. Then, it sets the fill color of the range **A2:C2**.

    Ranges are a fundamental part of Office Scripts in Excel. A range is a contiguous, rectangular block of cells that contains values, formula, and formatting. They are the basic structure of cells through which you'll perform most of your scripting tasks.

1. Add the following line to the end of the script (between where the `color` is set and the closing `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

1. Test the script by selecting **Run**. Your workbook should now look like this:

    :::image type="content" source="../images/tutorial-2.png" alt-text="A worksheet showing the fruit sales data row with the 'Oranges' row highlighted in the color orange and the 'Lemons' row highlighted in the color yellow.":::

## Create a table

Next, convert this fruit sales data into a table. You'll keep modifying the first script for the entire tutorial.

1. Add the following line to the end of the script (before the closing `}`):

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

1. That call returns a `Table` object. Use that table to sort the data. Sort the data in ascending order based on the values in the "Fruit" column. Add the following line after the table creation:

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    Your script should look like this:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    Tables have a `TableSort` object, accessed through the `Table.getSort` method. You can apply sorting criteria to that object. The `apply` method takes in an array of `SortField` objects. In this case, you only have one sorting criteria, so you only use one `SortField`. The `key: 0` value sets the column with the sort-defining values to "0" (which is the first column on the table, column **A** in this case). The `ascending: true` value sorts the data in ascending order (instead of descending order).

1. Run the script. You should see a table like this:

    :::image type="content" source="../images/tutorial-3.png" alt-text="A worksheet showing the sorted fruit sales table.":::

    > [!NOTE]
    > If you re-run the script, you'll get an error. This is because you can't create a table on top of another table. However, you can run the script in a different worksheet or workbook.

### Re-run the script

1. Create a new worksheet in the current workbook.
1. Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.
1. Run the script.

## Next steps

Complete [Tutorial: Clean and normalize Excel workbook data](excel-read-tutorial.md). It teaches you how to read data from a workbook with an Office Script.
