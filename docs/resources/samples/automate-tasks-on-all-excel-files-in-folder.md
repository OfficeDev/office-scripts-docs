---
title: 'Run a script on all Excel files in a folder'
description: 'Learn how to run a script on all the Excel files in a folder on OneDrive for Business.'
ms.date: 04/02/2021
localization_priority: Normal
---

# Run a script on all Excel files in a folder

This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business. It could also be used on a SharePoint folder.
It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.

Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!

## Sample code: Add formatting and insert comment

This is the script that runs on each individual workbook.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
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
1. Select **Manually trigger a flow** and press **Create**.
1. Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.
    ![The completed OneDrive for Business connector.](../../images/all-files-in-folder-sample-flow-1.png)
1. Select the "Sales" folder with the extracted workbooks.
1. To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:
    1. **Name** (the OneDrive file name value)
    1. "ends with"
    1. "xlsx".
    ![The apply-to-each file step with the condition.](../../images/all-files-in-folder-sample-flow-2.png)
1. Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action. Use the following values for the action:
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **File**: **Id** (the OneDrive file ID value)
    1. **Script**: Your script name
    ![The completed Excel Online (Business) connector.](../../images/all-files-in-folder-sample-flow-3.png)
1. Save the flow and try it out.

## Training video: Run a script on all Excel files in a folder

[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.
