---
title: 'Convert CSV files to Excel workbooks'
description: 'Learn how to use Office Scripts and Power Automate to create .xlsx files from .csv files.'
ms.date: 07/19/2021
localization_priority: Normal
---

# Convert CSV files to Excel workbooks

Many services export data as comma-separated value (CSV) files. This solution automates the process of converting those CSV files to Excel workbooks in the .xlsx file format. It uses a [Power Automate](https://flow.microsoft.com) flow to find files with the .csv extension in a OneDrive folder and an Office Script to copy the data from the .csv file into a new Excel workbook.

## Solution

1. Store the .csv files and a blank "Template" .xlsx file in a OneDrive folder.
1. Create an Office Script to parse the CSV data into a range.
1. Create a Power Automate flow to read the .csv files and pass their contents to the script.

## Sample files

Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> to get the Template.xlsx file and two sample .csv files. Extract the files into a folder in your OneDrive. This sample assumes the folder is named "output".

## Sample code: Insert comma-separated values into a workbook

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    data.push(value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g));
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## Power Automate flow: Create new .xlsx files

1. Sign into [Power Automate](https://flow.microsoft.com) and create a new **Scheduled cloud flow**.
1. Set the flow to **Repeat every** "1" "Day" and select **Create**.
1. Get the template Excel file. This is the basis for all the converted .csv files. Add a **New step** that uses the **OneDrive for Business** connector and the **Get file content** action. Provide the file path to the "Template.xlsx" file.
    * **File**: /output/Template.xlsx
1. Rename the **Get file content** step by going the **...** menu of that step (in the upper right corner of the connector) and selecting the **Rename** option. Change the step name to "Get Excel template".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="The completed OneDrive for Business connector in Power Automate, renamed to be Get Excel template.":::
1. Get all the files in the "output" folder. Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action. Provide the folder path that contains the .csv files.
    * **Folder**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="The completed OneDrive for Business connector in Power Automate.":::
1. Add a condition so that the flow only operates on .csv files. Add a **New step** that is the **Condition** control. Use the following values for the **Condition**.
    * **Choose a value**: *Name* (dynamic content from **List files in folder**). Note that this dynamic content has multiple results, so an **Apply to each** *value* control surrounds the **Condition**.
    * **ends with** (from the dropdown list)
    * **Choose a value**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="The completed Condition control with the Apply to each control around it.":::
1. The rest of the flow is under the **If yes** section, since we only want to act on .csv files. Get an individual .csv file by adding a **New step** that uses the **OneDrive for Business** connector and the **Get file content** action. Use the **Id** from the dynamic content from **List files in folder**.
    * **File**: *Id* (dynamic content from the **List files in folder** step)
1. Rename the new **Get file content** step to "Get .csv file". This helps distinguish this file from the Excel template.
1. Make the new .xlsx file, using the Excel template as the base content. Add a **New step** that uses the **OneDrive for Business** connector and the **Create file** action. Use the following values.
    * **Folder Path**: /output
    * **File Name**: *Name without extension*.xlsx (choose the *Name without extension* dynamic content from the **List files in folder** and manually type ".xlsx" after it)
    * **File Content**: *File content* (dynamic content from **Get Excel template**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="The Get .csv file and Create file steps of the Power Automate flow.":::
1. Run the script to copy data into the new workbook. Add the **Excel Online (Business)** connector with the **Run script** action. Use the following values for the action.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: *Id* (dynamic content from **Create file**)
    * **Script**: Convert CSV
    * **csv**: *File content* (dynamic content from **Get .csv file**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="The completed Excel Online (Business) connector in Power Automate.":::
