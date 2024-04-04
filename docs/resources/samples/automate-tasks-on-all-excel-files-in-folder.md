---
title: Run a script on all Excel files in a folder
description: Learn how to run a script on all the Excel files in a folder on OneDrive for Business.
ms.date: 11/30/2023
ms.localizationpriority: medium
---

# Run a script on all Excel files in a folder

This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business. It could also be used on a SharePoint folder.
It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.

## Sample Excel files

Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample. Extract those files to a folder titled **Sales**. Add the following script to your script collection to try the sample yourself!

## Sample code: Add formatting and insert comment

This is the script that runs on each individual workbook. In Excel,  **Automate** > **New Script** to paste the code and save the script. Save it as **Review script** and try the sample yourself!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  const table1 = workbook.getTable("Table1");

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
  highestAmountDue.getFormat().getFill().setColor("FFFF00");

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
1. In the flow builder, select the **+** button and **Add an action**. Use the **OneDrive for Business** connector's **List files in folder** action. Use the following values for the action.
    * **Folder**: /Sales (selected by the file picker)

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="The completed OneDrive for Business connector in Power Automate.":::

1. Ensure only workbooks are selected. Add a new **Condition** control action. Use the following values for the condition.
    * **Choose a value**: Name (_dynamic content from **List files in folder**_)
    * **ends with**: (from the dropdown list)
    * **Choose a value**: .xlsx

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="The Power Automate condition block that applies subsequent actions to each file.":::

1. Under the **True** branch, add a new action. Select the **Excel Online (Business)** connector's **Run script** action. Use the following values for the action.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: Id (_dynamic content from **List files in folder**_)
    * **Script**: Review script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="The completed Excel Online (Business) connector in Power Automate.":::

1. Save the flow. The flow designer should look like the following image.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-4.png" alt-text="A diagram of the completed flow that shows two steps before a condition and one step under the true path of the condition.":::

1. Try it out! Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.

## Training video: Run a script on all Excel files in a folder

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).
