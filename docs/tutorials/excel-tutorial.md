---
title: 'Record, edit, and create Office Scripts in Excel on the web'
description: 'A tutorial about the basics of Office Scripts, including recording scripts with the Action Recorder and writing data to a workbook.'
ms.date: 01/27/2020
localization_priority: Normal
---

# Record, edit, and create Office Scripts in Excel on the web

This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.

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

3. Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.
4. Press the **Record Actions** button.
5. Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.
6. Stop the recording by pressing the **Stop** button.
7. Fill in the **Script Name** field with a memorable name.
8. *Optional:* Fill in the **Description** field with a meaningful description. This is used to provide context as to what the script does. For the tutorial, you can use "Color-codes rows of a table".

   > [!TIP]
   > You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.

9. Save the script by pressing the **Save** button.

    Your worksheet should look like this (don't worry if the color is different):

    ![A fruit sales data row with the "Oranges" row highlighted orange.](../images/tutorial-1.png)

## Edit an existing script

The previous script colored the "Oranges" row to be orange. Let's add a yellow row for the "Lemons".

1. Open the **Automate** tab.
2. Press the **Code Editor** button.
3. Open the script you recorded in the previous section. You should see something similar to this code:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
    }
    ```

    This code gets the current worksheet by first accessing the workbook's worksheet collection. Then, it sets the fill color of the range **A2:C2**.

    Ranges are a fundamental part of Office Scripts in Excel on the web. A range is a contiguous, rectangular block of cells that contains values, formula, and formatting. They are the basic structure of cells through which you'll perform most of your scripting tasks.

4. Add the following line to the end of the script (between where the `color` is set and the closing `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
    ```

5. Test the script by pressing **Run**. Your workbook should now look like this:

    ![A fruit sales data row with the "Oranges" row highlighted orange and the "Lemons" row highlighted yellow.](../images/tutorial-2.png)

## Create a table

Let's convert this fruit sales data into a table. We'll use our script for the entire process.

1. Add the following line to the end of the script (before the closing `}`):

    ```TypeScript
    let table = selectedSheet.tables.add("A1:C5", true);
    ```

2. That call returns a `Table` object. Let's use that table to sort the data. We'll sort the data in ascending order based on the values in the "Fruit" column. Add the following line after the table creation:

    ```TypeScript
    table.sort.apply([{ key: 0, ascending: true }]);
    ```

    Your script should look like this:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
      selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
      let table = selectedSheet.tables.add("A1:C5", true);
      table.sort.apply([{ key: 0, ascending: true }]);
    }
    ```

    Tables have a `TableSort` object, accessed through the `Table.sort` property. You can apply sorting criteria to that object. The `apply` method takes in an array of `SortField` objects. In this case, we only have one sorting criteria, so we only use one `SortField`. `key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case). `ascending: true` sorts the data in ascending order (instead of descending order).

3. Run the script. You should see a table like this:

    ![A sorted fruit sales table.](../images/tutorial-3.png)

    > [!NOTE]
    > If you re-run the script, you'll get an error. This is because you cannot create a table on top of another table. However, you can run the script on a different worksheet or workbook.

### Re-run the script

1. Create a new worksheet in the current workbook.
2. Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.
3. Run the script.

## Next steps

Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial. It teaches you how to read data from a workbook with an Office Script.
