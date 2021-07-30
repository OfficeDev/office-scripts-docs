---
title: 'Run a script on all Excel files in a folder'
description: 'Learn how to run a script on all the Excel files in a folder on OneDrive for Business.'
ms.date: 06/29/2021
localization_priority: Normal
---

# Run a script on all Excel files in a folder

This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business. It could also be used on a SharePoint folder.
It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.

## Sample Excel files

Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample. Extract those files to a folder titled **Sales**. Add the following script to your script collection to try the sample yourself!

## Sample code: Add formatting and insert comment

This is the script that runs on each individual workbook.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## Power Automate flow: Run the script on every workbook in the folder

This flow runs the script on every workbook in the "Sales" folder.

1. Create a new **Instant cloud flow**.
1. Choose **Manually trigger a flow** and select **Create**.
1. Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="The completed OneDrive for Business connector in Power Automate.":::
1. Select the "Sales" folder with the extracted workbooks.
1. To ensure only workbooks are selected, choose **New step**, then select **Condition**. Use the following values for the action.
    1. **Name** (the OneDrive file name value)
    1. "ends with"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="The Power Automate condition block that applies subsequent actions to each file.":::
1. Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action. Use the following values for the action.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **File**: **Id** (the OneDrive file ID value)
    1. **Script**: Your script name

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="The completed Excel Online (Business) connector in Power Automate.":::
1. Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.

## Training video: Run a script on all Excel files in a folder

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).
