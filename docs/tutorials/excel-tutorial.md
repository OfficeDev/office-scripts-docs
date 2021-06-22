---
title: 'Record, edit, and create Office Scripts in Excel on the web'
description: 'A tutorial about the basics of Office Scripts, including recording scripts with the Action Recorder and writing data to a workbook.'
ms.date: 05/23/2021
localization_priority: Priority
---

# Record, edit, and create Office Scripts in Excel on the web

This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web. You'll record a script that applies some formatting to a sales record worksheet. You'll then edit the recorded script to apply more formatting, create a table, and sort that table. This record-then-edit pattern is an important tool to see what your Excel actions look like as code.

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.

## Add data and record a basic script

First, we'll need some data and a small starting script.

1. Create a new workbook in Excel for the Web.
2. Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.

    |Fruit |2018 |2019 |
    |:---|:---|:---|
    |Oranges |1000 |1200 |
    |Lemons |800 |900 |
    |Limes |600 |500 |
    |Grapefruits |900 |700 |

3. Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow. If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).
4. Press the **Record Actions** button.
5. Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.
6. Stop the recording by pressing the **Stop** button.

    Your worksheet should look like this (don't worry if the color is different):

    :::image type="content" source="../images/tutorial-1.png" alt-text="A worksheet showing fruit sales data row with the row containing 'Oranges' highlighted in the color orange.":::

## Edit an existing script

The previous script colored the "Oranges" row to be orange. Let's add a yellow row for the "Lemons".

1. From the now-open **Details** pane, press the **Edit** button.
2. You should see something similar to this code:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    This code gets the current worksheet from the workbook. Then, it sets the fill color of the range **A2:C2**.

    Ranges are a fundamental part of Office Scripts in Excel on the web. A range is a contiguous, rectangular block of cells that contains values, formula, and formatting. They are the basic structure of cells through which you'll perform most of your scripting tasks.

3. Add the following line to the end of the script (between where the `color` is set and the closing `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. Test the script by pressing **Run**. Your workbook should now look like this:

    :::image type="content" source="../images/tutorial-2.png" alt-text="A worksheet showing the fruit sales data row with the 'Oranges' row highlighted in the color orange and the 'Lemons' row highlighted in the color yellow.":::

## Create a table

Let's convert this fruit sales data into a table. We'll use our script for the entire process.

1. Add the following line to the end of the script (before the closing `}`):

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. That call returns a `Table` object. Let's use that table to sort the data. We'll sort the data in ascending order based on the values in the "Fruit" column. Add the following line after the table creation:

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

    Tables have a `TableSort` object, accessed through the `Table.getSort` method. You can apply sorting criteria to that object. The `apply` method takes in an array of `SortField` objects. In this case, we only have one sorting criteria, so we only use one `SortField`. `key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case). `ascending: true` sorts the data in ascending order (instead of descending order).

3. Run the script. You should see a table like this:

    :::image type="content" source="../images/tutorial-3.png" alt-text="A worksheet showing the sorted fruit sales table.":::

    > [!NOTE]
    > If you re-run the script, you'll get an error. This is because you cannot create a table on top of another table. However, you can run the script on a different worksheet or workbook.

### Re-run the script

1. Create a new worksheet in the current workbook.
2. Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.
3. Run the script.

## Next steps

Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial. It teaches you how to read data from a workbook with an Office Script.
