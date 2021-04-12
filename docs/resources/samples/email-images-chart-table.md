---
title: 'Email the images of an Excel chart and table'
description: 'Learn how to use Office Scripts and Power Automate to extract and email the images of an Excel chart and table.'
ms.date: 04/05/2021
localization_priority: Normal
---

# Use Office Scripts and Power Automate to email images of a chart and table

This sample uses Office Scripts and Power Automate to create a chart. It then emails images of the chart and its base table.

## Example scenario

* Calculate to get latest results.
* Create chart.
* Get chart and table images.
* Email the images with Power Automate.

_Input data_

![Input data](../../images/input-data.png)

_Output chart_

![Chart created](../../images/chart-created.png)

_Email that was received through Power Automate flow_

![Email received](../../images/email-received.png)

## Solution

This solution has two parts:

1. [An Office Script to calculate and extract Excel chart and table](#sample-code-calculate-and-extract-excel-chart-and-table)
1. A Power Automate flow to invoke the script and email the results. For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## Sample code: Calculate and extract Excel chart and table

The following script calculates and extracts an Excel chart and table.

Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## Power Automate flow: Email the chart and table images

This flow runs the script and emails the returned images.

1. Create a new **Instant cloud flow**.
1. Select **Manually trigger a flow** and press **Create**.
1. Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script (preview)** action. Use the following values for the action:
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Your script name

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="The completed Excel Online (Business) connector in Power Automate.":::
1. This sample uses Outlook as the email client. You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook. Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action. Use the following values for the action:
    * **To**: Your test email account (or personal email)
    * **Subject**: Please Review Report Data
    * For the **Body** field, select "Code View" (`</>`) and enter the following:

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="The completed Office 365 Outlook connector in Power Automate.":::
1. Save the flow and try it out.

## Training video: Extract and email images of chart and table

[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")
