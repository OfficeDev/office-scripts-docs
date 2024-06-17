---
title: Combine workbooks into a single workbook
description: Learn how to use Office Scripts and Power Automate to create merge worksheets from other workbooks into a single workbook.
ms.date: 11/29/2023
ms.localizationpriority: medium
---

# Combine worksheets into a single workbook

This sample shows how to pull data from multiple workbooks into a single, centralized workbook. It uses two scripts: one to retrieve information from a workbook and another to create new worksheets with that information. It combines the scripts in a Power Automate flow that acts on an entire OneDrive folder.

> [!IMPORTANT]
> This sample only copies the values from the other workbooks. It does not preserve formatting, charts, tables, or other objects.

## Solution

1. Create a new Excel file in your OneDrive. The file name "Combination.xlsx" is used in this sample.
1. Create and save the two scripts from this sample.
1. Create a folder in your OneDrive and add one or more workbooks with data to it. The folder name "output" is used in this sample.
1. Build a flow (as described in the [Power Automate flow](#power-automate-flow-combine-worksheets-into-a-single-workbook) section of this article) to perform these steps:
    1. List all the files the "output" folder.
    1. Use the **Return worksheet data** script to get the data from every worksheet in each of the workbooks.
    1. Use the **Add worksheets** script to create a new worksheet in the "Combination.xlsx" workbook for every worksheet in all the other files.

## Sample code: Return worksheet data

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[] {
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## Sample code: Add worksheets

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[]) {
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## Power Automate flow: Combine worksheets into a single workbook

1. Sign into [Power Automate](https://make.powerautomate.com/create) and create a new **Instant cloud flow**.
1. Choose **Manually trigger a flow** and select **Create**.
1. Get all the workbooks you want to combine from their folder. Add an action and choose the **OneDrive for Business** connector's **List files in folder** action. For the **Folder** field, use the file picker to select the "output" folder.

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="The completed OneDrive for Business connector in Power Automate.":::

1. Add an action to run the **Return worksheet data** script to get all the data from each of the workbooks. Choose the **Excel Online (Business)** connector's **Run script** action. Use the following values for the action. Note that when you add the *Id* for the file, Power Automate will wrap the action in a **For each** control, so the action will be performed on every file.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: *Id* (dynamic content from **List files in folder**)
    * **Script**: Return worksheet data

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="The completed Run script action in the action task pane.":::

1. Add an action to run the **Add worksheets** script on the new Excel file you created. This will add the data from all the other workbooks. After the previous **Run script** action and inside the **For each** control, add an action that uses the **Excel Online (Business)** connector's **Run script** action. Use the following values for the action.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: "Combination.xlsx" (your file, as selected by the file picker)
    * **Script**: Add worksheets
    * **workbookName**: *Name* (dynamic content from **List files in folder**)
    * **worksheetInformation** (see the note following the next image): *result* (dynamic content from **Run script**)

    :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="The second Run script action inside the Apply to each control.":::
    > [!NOTE]
    > Select the **Switch to input entire array** button to add the array object directly, instead of individual items for the array. Do this before entering *result*.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-4.png" alt-text="The button to switch to input an entire array in a control field input box.":::

1. Save the flow. The flow designer should look like the following image.

    :::image type="content" source="../../images/combine-worksheets-flow-5.png" alt-text="The flow designer showing the two Run script actions inside a For each control loop.":::
  
1. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.
1. The "Combination.xlsx" file should now have new worksheets.

## Troubleshooting

- **A resource with the same name or identifier already exists**: This error likely indicates the "Combination.xlsx" workbook already has a worksheet with the same name. This will happen if you run the flow multiple times with the same workbooks. Create a new workbook each time to store the combined data or use different file names in the "output" folder.
- **The argument is invalid or missing or has an incorrect format**: This error can mean that the generated  worksheet name doesn't meet [Excel's requirements](https://support.microsoft.com/office/3f1f7148-ee83-404d-8ef0-9ff99fbad1f9). This is likely because the name is too long. If the worksheet names will be more than 30 characters replace the code in "Add worksheets" that calls `addWorksheet` with something that shortens the string. Since the workbook name itself might be too long, add an incrementing number to the end of the worksheet name. Declare this number outside of the `forEach` loop.

  ```TypeScript
  let worksheetNumber = 1;
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
      let worksheetName = `${workbookName}.${value.name}`;
      let sheet = workbook.addWorksheet(`${worksheetName.substr(0,30)}${worksheetNumber++}`);
  ```

  Additionally, if the workbook names are longer than 30 characters, you'll need to shorten them in the flow. First, you must create a variable in the flow to track the workbook count. This will avoid identical shortened names being passed to the script. Add an **Initialize variable** action before the flow (of **Type** "Integer") and an **Increment variable** action between the two **Run script** actions. Then, instead of using *Name* as the **workbookName** in "Run script 1", use the expression `substring(items('Apply_to_each')?['Name'],0,min(length(items('Apply_to_each')?['Name']),20))` and the dynamic content from your variable. This shortens the workbook names to 20 characters and appends the current workbook number to the string being passed to the script.

  :::image type="content" source="../../images/combine-worksheets-flow-workbook-name-shorten.png" alt-text="The second Run script action with the changes to the workbook name parameter..":::

  :::image type="content" source="../../images/combine-worksheets-flow-workbook-name-shorten-flow.png" alt-text="The Initialize variable and Increment variable steps added to the flow.":::

  > [!NOTE]
  > Rather than making the flow and script more complicated, it might be easier to guarantee the file and worksheet names are short enough.
