---
title: 'Integrate Office Scripts into automated Power Automate flows'
description: 'A tutorial about integrating Power Automate with Office Scripts for Excel on the web using automatic external triggers, such as receiving mail through Outlook.'
ms.date: 01/29/2020
localization_priority: Normal
---

# Integrate Office Scripts into automated Power Automate flows (preview)

This tutorial will teach you use an Office Script for Excel on the web in an automated [Power Automate](https://flow.microsoft.com) workflow.

## Prerequisites

[!INCLUDE [Preview note](../includes/preview-note.md)]

Before starting this tutorial, you'll need access to Office Scripts, which requires the following:

- [Excel on the web](https://www.office.com/launch/excel).
- Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon in Excel on the web.
- [Preview access to Power Automate](https://flow.microsoft.com/blog/flow-preview-program/).

> [!IMPORTANT]
> This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.

## Prepare the workbook

Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components. So, we need a workbook and worksheet with consistent names that Power Automate can reference.

1. Create a new workbook named **MyWorkbook**.

2. Go to the **Automate** tab and select **Code Editor**.

3. Select **New Script**.

4. Run the following script to setup the workbook with consistent worksheet, table, and PivotTable names.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Add a new worksheet to store our email table
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let emailsSheet = worksheets.add("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").values = [
        ["Date", "Day of the week", "Email address", "Subject"]
      ];
      let tables = workbook.tables;
      let newTable = tables.add(emailsSheet.getRange("A1:D2"), true);

      // Add a new pivot table to a new worksheet
      let pivotWorksheet = worksheets.add();
      let pivotTables = workbook.pivotTables;
      let newPivotTable = pivotTables.add("PivotTable3", "Table6", pivotWorksheet.getRange("A3:C20"));
      pivotWorksheet.name = "Pivot";

      // Setup the pivot hierarchies
      newPivotTable.rowHierarchies.add(newPivotTable.hierarchies.getItem("Day of the week"));
      newPivotTable.rowHierarchies.add(newPivotTable.hierarchies.getItem("Email address"));
      newPivotTable.dataHierarchies.add(newPivotTable.hierarchies.getItem("Subject"));
    }
    ```

## Create an Office Script for your automated workflow

1. From within the **Code Editor**, select **New Script**.


## Create an automated workflow with Power Automate

1. Sign in to the [Power Automate preview site](https://preview.flow.microsoft.com).

2. In the menu that's displayed on the left side of the screen, press **Create**. This brings you to list of ways to create new workflows.

    ![The Create button in Power Automate.](../images/power-automate-tutorial-1.png)

## Run the script through Power Automate

1. From the main Power Automate page, select **My flows**.

    ![The My flows button in Power Automate.](../images/power-automate-tutorial-7.png)

## Next steps

Visit the [Power Automate Documentation](https://docs.microsoft.com/power-automate) to learn ways to automate your Office Scripts.
