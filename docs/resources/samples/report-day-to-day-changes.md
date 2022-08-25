---
title: Record day-to-day changes in a workbook and report them with a flow
description: Learn how to use Office Scripts and Power Automate to track value changes in a workbook
ms.date: 08/25/2022
ms.localizationpriority: medium
---

# Record day-to-day changes in a workbook and report them with a flow

Power Automate and Office Scripts combine to handle repetitive tasks for you. In this sample, you're tasked with recording a single numerical reading in a workbook every day and reporting the change since yesterday. You'll build a flow to get that reading, log it in the workbook, and report the change through an email.

## Sample Excel file

Download <a href="daily-readings.xlsx">daily-readings.xlsx</a> for a ready-to-use workbook. Add the following script to try the sample yourself!

## Sample code: Record and report daily readings

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## Sample flow: Report day-to-day changes

Follow these steps to build a [Power Automate](https://powerautomate.microsoft.com/) flow for the sample.

1. Create a new **Schedule cloud flow**.
1. Schedule the flow to repeat every **1 Day**.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="The flow creation step showing it will repeat every day.":::
1. Select **Create**.
1. In a real flow, you would add a step that gets your data. This could be from another workbook, a Teams adaptive card, or any other source. To test the sample, make a test number. Add a new step with the **Initialize variable** action. Give it the following values.
    1. **Name**: Input
    1. **Type**: Integer
    1. **Value**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="The Initialize variable action with the given values.":::
1. Add a new step with the **Excel Online (Business)** connector with the **Run script** action. Use the following values for the action.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **File**: daily-readings.xlsx *(Chosen through the file browser)*
    1. **Script**: Your script name
    1. **newData**: Input *(dynamic content)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="The Run script action with the given values.":::
1. The script returns the daily reading difference as dynamic content named "result". As an example, you an email the information to yourself. Create a new step that uses the **Outlook** connector with the **Send an email (V2)** action (or whatever email client you prefer). Use the following values to complete the action.
    1. **To**: Your email address
    1. **Subject**: Daily reading change
    1. **Body**: "Difference from yesterday" result *(dynamic content from Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="The completed Outlook connector in Power Automate.":::
1. Save the flow and try it out. Use the **Test** button on the flow editor page. Be sure to allow access when prompted.
